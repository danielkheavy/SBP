VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form tfactura 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Documentos Valorados"
   ClientHeight    =   9060
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   18900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   18900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdoData 
      Height          =   375
      Left            =   15000
      Top             =   2880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Password=mastercard;Persist Security Info=True;User ID=sa;Initial Catalog=FEGRIFO"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=mastercard;Persist Security Info=True;User ID=sa;Initial Catalog=FEGRIFO"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "detalle"
      Caption         =   "Data"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox estado_sunat 
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
      Left            =   14040
      MaxLength       =   50
      TabIndex        =   236
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
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
      Height          =   9015
      Left            =   13920
      TabIndex        =   193
      Top             =   4080
      Visible         =   0   'False
      Width           =   14775
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Habilitar Proveedor"
         Height          =   375
         Left            =   7200
         TabIndex        =   199
         Top             =   225
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox buffer1 
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
         Left            =   2565
         MaxLength       =   10
         TabIndex        =   198
         TabStop         =   0   'False
         Top             =   555
         Width           =   2895
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
         Left            =   2550
         MaxLength       =   10
         TabIndex        =   197
         TabStop         =   0   'False
         Top             =   225
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
         Left            =   105
         Style           =   2  'Dropdown List
         TabIndex        =   196
         TabStop         =   0   'False
         Top             =   300
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
         Left            =   5400
         TabIndex        =   195
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   194
         TabStop         =   0   'False
         Top             =   600
         Width           =   2295
      End
      Begin MSDataGridLib.DataGrid dbgrid89 
         Height          =   855
         Left            =   120
         TabIndex        =   200
         Top             =   7800
         Visible         =   0   'False
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   1508
         _Version        =   393216
         HeadLines       =   1
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
               LCID            =   10250
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
               LCID            =   10250
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
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   6735
         Left            =   120
         TabIndex        =   201
         Top             =   960
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   11880
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
      Begin MSDataGridLib.DataGrid dbgrid3 
         Height          =   6735
         Left            =   120
         TabIndex        =   202
         Top             =   960
         Visible         =   0   'False
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   11880
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648447
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
   Begin VB.Frame Frame6 
      BackColor       =   &H00808080&
      Caption         =   "Control de Peso"
      Height          =   7485
      Left            =   3240
      TabIndex        =   216
      Top             =   9120
      Visible         =   0   'False
      Width           =   10440
      Begin VB.CommandButton Command15 
         Caption         =   "Copiar"
         Height          =   615
         Left            =   9105
         TabIndex        =   232
         Top             =   1275
         Width           =   1215
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Close"
         Height          =   615
         Left            =   9165
         TabIndex        =   231
         Top             =   510
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid dbgrid13 
         Height          =   4935
         Left            =   240
         TabIndex        =   217
         Top             =   360
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   8705
         _Version        =   393216
         HeadLines       =   2
         RowHeight       =   29
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "NroJabas"
            Caption         =   "NroJabas"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "CantProd"
            Caption         =   "CantProd"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "PesoBruto"
            Caption         =   "PesoBruto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Tara"
            Caption         =   "Tara"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "PesoNeto"
            Caption         =   "PesoNeto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Precio"
            Caption         =   "Precio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Total"
            Caption         =   "Total"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "Producto"
            Caption         =   "Producto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
         EndProperty
      End
      Begin VB.Label nsdx6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   7440
         TabIndex        =   230
         Top             =   6120
         Width           =   1575
      End
      Begin VB.Label Label67 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   6240
         TabIndex        =   229
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label nsdx5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   5040
         TabIndex        =   228
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label nsdx4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   3840
         TabIndex        =   227
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label nsdx3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   2640
         TabIndex        =   226
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label nsdx2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   1440
         TabIndex        =   225
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label nsdx1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   240
         TabIndex        =   224
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio"
         Height          =   495
         Left            =   6240
         TabIndex        =   223
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Label Label64 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PesoNeto"
         Height          =   495
         Left            =   5040
         TabIndex        =   222
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Label Label63 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tara"
         Height          =   495
         Left            =   3840
         TabIndex        =   221
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Label Label62 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PesoBruto"
         Height          =   495
         Left            =   2640
         TabIndex        =   220
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Label Label61 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TotProducto"
         Height          =   495
         Left            =   1440
         TabIndex        =   219
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Label Label60 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TotJabas"
         Height          =   495
         Left            =   240
         TabIndex        =   218
         Top             =   5640
         Width           =   1215
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Cargar Saldo a una fecha Determinado"
      Height          =   7575
      Left            =   5760
      TabIndex        =   208
      Top             =   8940
      Visible         =   0   'False
      Width           =   10095
      Begin VB.TextBox fechai 
         Height          =   615
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   212
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox fechaf 
         Height          =   615
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   211
         Top             =   960
         Width           =   2655
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5775
         Left            =   120
         TabIndex        =   209
         Top             =   1680
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   10186
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   29
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
               LCID            =   10250
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
               LCID            =   10250
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
      Begin VB.PictureBox EC_Button1 
         BackColor       =   &H00808080&
         Height          =   735
         Left            =   4800
         ScaleHeight     =   675
         ScaleWidth      =   1395
         TabIndex        =   210
         Top             =   360
         Width           =   1455
      End
      Begin VB.PictureBox EC_Button2 
         BackColor       =   &H00808080&
         Height          =   735
         Left            =   7440
         ScaleHeight     =   675
         ScaleWidth      =   1395
         TabIndex        =   213
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label53 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"TFACTURA.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   135
         TabIndex        =   215
         Top             =   390
         Width           =   1935
      End
      Begin VB.Label Label41 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   7440
         TabIndex        =   214
         Top             =   2640
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Lista Precios"
      Height          =   4815
      Left            =   5730
      TabIndex        =   203
      Top             =   9255
      Visible         =   0   'False
      Width           =   8295
      Begin VB.CommandButton Command8 
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
         Left            =   7440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TFACTURA.frx":009A
         Style           =   1  'Graphical
         TabIndex        =   205
         ToolTipText     =   "Borrar registro"
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   204
         Top             =   360
         Width           =   3375
      End
      Begin MSDBGrid.DBGrid DBGrid4 
         Height          =   3615
         Left            =   120
         OleObjectBlob   =   "TFACTURA.frx":12AC
         TabIndex        =   206
         Top             =   960
         Width           =   7215
      End
      Begin VB.Label tproducto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
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
         ForeColor       =   &H00FFFFC0&
         Height          =   420
         Left            =   3600
         TabIndex        =   207
         Top             =   360
         Width           =   135
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00808080&
      Caption         =   "Datos Adicionales"
      Height          =   5895
      Left            =   2205
      TabIndex        =   174
      Top             =   8010
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox destino 
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
         Left            =   1800
         MaxLength       =   60
         TabIndex        =   184
         TabStop         =   0   'False
         Top             =   3360
         Width           =   3375
      End
      Begin VB.TextBox partida 
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
         Left            =   1800
         MaxLength       =   60
         TabIndex        =   183
         TabStop         =   0   'False
         Top             =   3000
         Width           =   3375
      End
      Begin VB.TextBox hora 
         Enabled         =   0   'False
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
         TabIndex        =   182
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox servicio 
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
         MaxLength       =   1
         TabIndex        =   181
         TabStop         =   0   'False
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox cajero 
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
         MaxLength       =   13
         TabIndex        =   180
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox turno 
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
         MaxLength       =   1
         TabIndex        =   179
         TabStop         =   0   'False
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox caja 
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
         MaxLength       =   2
         TabIndex        =   178
         TabStop         =   0   'False
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command11 
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
         Left            =   6360
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TFACTURA.frx":230F
         Style           =   1  'Graphical
         TabIndex        =   177
         TabStop         =   0   'False
         ToolTipText     =   "Grabar registro"
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command10 
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
         Left            =   6360
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TFACTURA.frx":3521
         Style           =   1  'Graphical
         TabIndex        =   176
         ToolTipText     =   "Borrar registro"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox fechasunat 
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
         TabIndex        =   175
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dir.Destino"
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
         Left            =   120
         TabIndex        =   192
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dir.Partida"
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
         Left            =   120
         TabIndex        =   191
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label54 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hora"
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
         Left            =   120
         TabIndex        =   190
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label43 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "servicio  (D C A)"
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
         Left            =   120
         TabIndex        =   189
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label42 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cajero"
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
         Left            =   120
         TabIndex        =   188
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label40 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Turno"
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
         Left            =   120
         TabIndex        =   187
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label39 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja"
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
         Left            =   120
         TabIndex        =   186
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label27 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Sunat"
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
         Left            =   120
         TabIndex        =   185
         Top             =   1680
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Caption         =   "Observaciones"
      Height          =   3855
      Left            =   1350
      TabIndex        =   154
      Top             =   8940
      Visible         =   0   'False
      Width           =   6975
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
         Left            =   5760
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TFACTURA.frx":4733
         Style           =   1  'Graphical
         TabIndex        =   160
         ToolTipText     =   "Grabar registro"
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command5 
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
         Left            =   5760
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TFACTURA.frx":5945
         Style           =   1  'Graphical
         TabIndex        =   159
         ToolTipText     =   "Borrar registro"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox observa4 
         Height          =   375
         Left            =   120
         MaxLength       =   40
         TabIndex        =   158
         TabStop         =   0   'False
         Top             =   1800
         Width           =   5055
      End
      Begin VB.TextBox observa3 
         Height          =   375
         Left            =   120
         MaxLength       =   40
         TabIndex        =   157
         TabStop         =   0   'False
         Top             =   1440
         Width           =   5055
      End
      Begin VB.TextBox observa2 
         Height          =   375
         Left            =   120
         MaxLength       =   40
         TabIndex        =   156
         TabStop         =   0   'False
         Top             =   1080
         Width           =   5055
      End
      Begin VB.TextBox observa1 
         Height          =   375
         Left            =   120
         MaxLength       =   40
         TabIndex        =   155
         TabStop         =   0   'False
         Top             =   720
         Width           =   5055
      End
   End
   Begin VB.TextBox localf 
      Enabled         =   0   'False
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
      Left            =   13080
      MaxLength       =   11
      TabIndex        =   136
      TabStop         =   0   'False
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox local1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   10680
      MaxLength       =   2
      TabIndex        =   14
      Top             =   1200
      Width           =   975
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "CodigoProveedor"
      Height          =   1935
      Left            =   5640
      TabIndex        =   120
      Top             =   5160
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox producto 
         Height          =   495
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   125
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command13 
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
         Left            =   3480
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TFACTURA.frx":6B57
         Style           =   1  'Graphical
         TabIndex        =   124
         TabStop         =   0   'False
         ToolTipText     =   "Grabar registro"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox rcodigo 
         Height          =   495
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   122
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Command12 
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
         Left            =   3480
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TFACTURA.frx":7D69
         Style           =   1  'Graphical
         TabIndex        =   121
         TabStop         =   0   'False
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label33 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto"
         Height          =   495
         Left            =   120
         TabIndex        =   126
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label32 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo Proveedor"
         Height          =   495
         Left            =   120
         TabIndex        =   123
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   18840
      TabIndex        =   111
      Top             =   0
      Width           =   18900
      Begin VB.CommandButton cmdGuiaRemision 
         Caption         =   "Complementar Guia Electronica de Remision"
         Height          =   555
         Left            =   12735
         TabIndex        =   243
         Top             =   60
         Width           =   2025
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
         Left            =   1440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TFACTURA.frx":8F7B
         Style           =   1  'Graphical
         TabIndex        =   114
         TabStop         =   0   'False
         ToolTipText     =   "Salir"
         Top             =   0
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
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TFACTURA.frx":A18D
         Style           =   1  'Graphical
         TabIndex        =   113
         TabStop         =   0   'False
         ToolTipText     =   "Grabar registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdAddEntry 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
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
         Picture         =   "TFACTURA.frx":B39F
         Style           =   1  'Graphical
         TabIndex        =   112
         TabStop         =   0   'False
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label znumero 
         Height          =   375
         Left            =   10035
         TabIndex        =   118
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label zserie 
         Height          =   375
         Left            =   9195
         TabIndex        =   117
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label ztipo 
         Height          =   375
         Left            =   8475
         TabIndex        =   116
         Top             =   120
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label bandera 
         BackColor       =   &H00E0E0E0&
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
         Height          =   375
         Left            =   8760
         TabIndex        =   115
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Cargar Productos "
      Height          =   2775
      Left            =   5520
      TabIndex        =   100
      Top             =   4920
      Visible         =   0   'False
      Width           =   5775
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
         Height          =   615
         Left            =   4800
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TFACTURA.frx":C5B1
         Style           =   1  'Graphical
         TabIndex        =   104
         ToolTipText     =   "Borrar registro"
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command7 
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
         Left            =   4800
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TFACTURA.frx":D7C3
         Style           =   1  'Graphical
         TabIndex        =   103
         ToolTipText     =   "Grabar registro"
         Top             =   1080
         Width           =   735
      End
      Begin VB.ComboBox Combo3 
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   102
         Top             =   840
         Width           =   3495
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Chequear dia de Visita"
         Height          =   375
         Left            =   240
         TabIndex        =   101
         Top             =   1200
         Width           =   3495
      End
      Begin VB.Label Label24 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Carga"
         Height          =   375
         Left            =   240
         TabIndex        =   105
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Ingreso de Lineas"
      Height          =   3255
      Left            =   585
      TabIndex        =   58
      Top             =   8760
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
         Picture         =   "TFACTURA.frx":E9D5
         Style           =   1  'Graphical
         TabIndex        =   76
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
         Picture         =   "TFACTURA.frx":FBE7
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Grabar registro"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox t16 
         Height          =   375
         Left            =   6240
         MaxLength       =   2
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t15 
         Height          =   375
         Left            =   6240
         MaxLength       =   2
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t14 
         Height          =   375
         Left            =   6240
         MaxLength       =   2
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox t13 
         Height          =   375
         Left            =   6240
         MaxLength       =   2
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t12 
         Height          =   375
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t11 
         Height          =   375
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t10 
         Height          =   375
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox t9 
         Height          =   375
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t8 
         Height          =   375
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t7 
         Height          =   375
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t6 
         Height          =   375
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox t5 
         Height          =   375
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t4 
         Height          =   375
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t3 
         Height          =   375
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t2 
         Height          =   375
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox t1 
         Height          =   375
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.Label linea 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4440
         TabIndex        =   99
         Top             =   360
         Width           =   855
      End
      Begin VB.Label nt16 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   98
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label nt15 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   97
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt14 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   96
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label nt13 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   95
         Top             =   840
         Width           =   735
      End
      Begin VB.Label nt12 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   94
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label nt11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   93
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   92
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label nt9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   91
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label30 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   90
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label29 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   89
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label28 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   88
         Top             =   840
         Width           =   735
      End
      Begin VB.Label nt8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   87
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label nt7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   86
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   85
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label nt5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   84
         Top             =   840
         Width           =   735
      End
      Begin VB.Label nt4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   83
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label nt3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   82
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   81
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label nt1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   80
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tallas"
         Height          =   375
         Left            =   240
         TabIndex        =   79
         Top             =   840
         Width           =   975
      End
      Begin VB.Label nlinea 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   78
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Linea"
         Height          =   375
         Left            =   240
         TabIndex        =   77
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox moneda 
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
      Left            =   4560
      MaxLength       =   1
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1560
      Width           =   615
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   15120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.TextBox fpago 
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
      Left            =   4560
      MaxLength       =   3
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox observa 
      BackColor       =   &H00FFFFFF&
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
      Left            =   10680
      MaxLength       =   60
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1920
      Width           =   4095
   End
   Begin VB.TextBox bodegaf 
      Enabled         =   0   'False
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
      Left            =   13080
      MaxLength       =   2
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox bodega 
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
      Left            =   10680
      MaxLength       =   2
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox dias 
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
      Left            =   7440
      MaxLength       =   10
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox paridad 
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
      Left            =   7440
      MaxLength       =   10
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox transporte 
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
      Left            =   7440
      MaxLength       =   11
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox vendedor 
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
      Left            =   7440
      MaxLength       =   5
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox fechae 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4560
      MaxLength       =   10
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox fecha 
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
      Left            =   4560
      MaxLength       =   10
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox codigo 
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
      Left            =   840
      MaxLength       =   13
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox ttipo 
      BackColor       =   &H00C0FFFF&
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
      Left            =   840
      MaxLength       =   3
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox numero 
      BackColor       =   &H00C0FFFF&
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
      Left            =   840
      MaxLength       =   25
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2400
   End
   Begin VB.CheckBox saldoini 
      BackColor       =   &H00E0E0E0&
      Caption         =   "SdoInicial"
      Height          =   255
      Left            =   5280
      TabIndex        =   153
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox sinigv 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sin Igv"
      Height          =   255
      Left            =   5280
      TabIndex        =   142
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CheckBox dflag 
      BackColor       =   &H00E0E0E0&
      Caption         =   "SinStock"
      Height          =   255
      Left            =   5280
      TabIndex        =   173
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox TIPONCD 
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
      Left            =   10680
      MaxLength       =   2
      TabIndex        =   234
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox serie 
      BackColor       =   &H00C0FFFF&
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
      Left            =   840
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1200
      Width           =   735
   End
   Begin VB.Frame documentorelacionado 
      BackColor       =   &H00808080&
      Caption         =   "        Documento relacionado     "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1180
      Left            =   11640
      TabIndex        =   238
      Top             =   750
      Visible         =   0   'False
      Width           =   3165
      Begin VB.TextBox serie11 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1160
         MaxLength       =   4
         TabIndex        =   239
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox numero11 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1750
         MaxLength       =   8
         TabIndex        =   241
         Top             =   600
         Width           =   1340
      End
      Begin VB.TextBox tipo11 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1160
         MaxLength       =   3
         TabIndex        =   237
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label48 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
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
         Left            =   70
         TabIndex        =   242
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label47 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie/Numero"
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
         Left            =   70
         TabIndex        =   240
         Top             =   600
         Width           =   1095
      End
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "TFACTURA.frx":10DF9
      Height          =   5055
      Left            =   0
      OleObjectBlob   =   "TFACTURA.frx":10E0D
      TabIndex        =   244
      TabStop         =   0   'False
      Top             =   2280
      Width           =   14775
   End
   Begin VB.Label Label46 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Serie"
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
      Left            =   0
      TabIndex        =   233
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label txestado 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4335
      TabIndex        =   172
      Top             =   9000
      Width           =   1575
   End
   Begin VB.Label Label59 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estado"
      Height          =   375
      Left            =   3240
      TabIndex        =   171
      Top             =   9000
      Width           =   1095
   End
   Begin VB.Label txdetraccion 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   13080
      TabIndex        =   170
      Top             =   9000
      Width           =   1695
   End
   Begin VB.Label txisc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   10080
      TabIndex        =   169
      Top             =   8640
      Width           =   1455
   End
   Begin VB.Label txivap 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   10080
      TabIndex        =   168
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label Label58 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Impuesto"
      Height          =   375
      Left            =   11640
      TabIndex        =   167
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label Label44 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Detraccion"
      Height          =   375
      Left            =   11640
      TabIndex        =   166
      Top             =   9000
      Width           =   1455
   End
   Begin VB.Label Label26 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Isc"
      Height          =   375
      Left            =   8760
      TabIndex        =   165
      Top             =   8640
      Width           =   1335
   End
   Begin VB.Label Label25 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ivap"
      Height          =   375
      Left            =   8760
      TabIndex        =   164
      Top             =   8280
      Width           =   1335
   End
   Begin VB.Label Label23 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Gravado"
      Height          =   375
      Left            =   6240
      TabIndex        =   163
      Top             =   9000
      Width           =   1095
   End
   Begin VB.Label Label22 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Subtotal"
      Height          =   375
      Left            =   6240
      TabIndex        =   162
      Top             =   8640
      Width           =   1095
   End
   Begin VB.Label Label20 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Neto"
      Height          =   375
      Left            =   6240
      TabIndex        =   161
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Label Label49 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Recalcular"
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
      TabIndex        =   152
      Top             =   9000
      Width           =   2775
   End
   Begin VB.Label importacion 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   15600
      TabIndex        =   151
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label16 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Gastos"
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
      Left            =   1440
      TabIndex        =   150
      Top             =   8640
      Width           =   1695
   End
   Begin VB.Label Label55 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Diferencia"
      Height          =   375
      Left            =   3240
      TabIndex        =   149
      Top             =   8640
      Width           =   1095
   End
   Begin VB.Label diferencia 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4320
      TabIndex        =   148
      Top             =   8640
      Width           =   1455
   End
   Begin VB.Label costofactura 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "precio"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   4320
      TabIndex        =   147
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Label Label52 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CostoFactura"
      Height          =   375
      Left            =   3240
      TabIndex        =   146
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Label proveedorp 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "proveedorp"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   4320
      TabIndex        =   145
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label Label50 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CostoAnterior"
      Height          =   375
      Left            =   3240
      TabIndex        =   144
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Label precio 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   0
      TabIndex        =   143
      Top             =   8640
      Width           =   1455
   End
   Begin VB.Label numeroimp 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   14880
      TabIndex        =   141
      Top             =   4440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label tipoimp 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   14880
      TabIndex        =   140
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label serieimp 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   14880
      TabIndex        =   139
      Top             =   4080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label costopais 
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   138
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Label Label38 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AlmacenDestino"
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
      Left            =   11640
      TabIndex        =   137
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label zlocal 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   18000
      TabIndex        =   135
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label37 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo"
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
      Left            =   0
      TabIndex        =   134
      Top             =   840
      Width           =   855
   End
   Begin VB.Label tflete 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   10080
      TabIndex        =   133
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label Label35 
      BackColor       =   &H00E0E0E0&
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
      Left            =   8760
      TabIndex        =   132
      Top             =   7920
      Width           =   1335
   End
   Begin VB.Label Label36 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
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
      Left            =   11640
      TabIndex        =   131
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label txpercepcio 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   13080
      TabIndex        =   130
      Top             =   8640
      Width           =   1695
   End
   Begin VB.Label Label34 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Percepcion"
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
      Left            =   11640
      TabIndex        =   129
      Top             =   8640
      Width           =   1455
   End
   Begin VB.Label nbodega1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      Height          =   195
      Left            =   13920
      TabIndex        =   128
      Top             =   2640
      Width           =   45
   End
   Begin VB.Label escompra 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1200
      TabIndex        =   127
      Top             =   9000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label cargado 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   14760
      TabIndex        =   119
      Top             =   7200
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label gravado 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7320
      TabIndex        =   110
      Top             =   9000
      Width           =   1455
   End
   Begin VB.Label Label31 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descuento"
      Height          =   375
      Left            =   6240
      TabIndex        =   109
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Label zona 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2520
      TabIndex        =   108
      Top             =   9000
      Width           =   600
   End
   Begin VB.Label racu 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14400
      TabIndex        =   107
      Top             =   0
      Width           =   255
   End
   Begin VB.Label acu1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14400
      TabIndex        =   106
      Top             =   240
      Width           =   255
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   360
      Left            =   2640
      Picture         =   "TFACTURA.frx":17A54
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   600
   End
   Begin VB.Label numero7 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   57
      Top             =   10320
      Width           =   1215
   End
   Begin VB.Label serie7 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   56
      Top             =   10320
      Width           =   735
   End
   Begin VB.Label numero6 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13560
      TabIndex        =   55
      Top             =   9600
      Width           =   1215
   End
   Begin VB.Label serie6 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12840
      TabIndex        =   54
      Top             =   10080
      Width           =   735
   End
   Begin VB.Label numero5 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11640
      TabIndex        =   53
      Top             =   10080
      Width           =   1215
   End
   Begin VB.Label serie5 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10920
      TabIndex        =   52
      Top             =   10080
      Width           =   735
   End
   Begin VB.Label numero4 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   51
      Top             =   10080
      Width           =   1215
   End
   Begin VB.Label serie4 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   50
      Top             =   10080
      Width           =   735
   End
   Begin VB.Label numero3 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13560
      TabIndex        =   49
      Top             =   9360
      Width           =   1215
   End
   Begin VB.Label serie3 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12840
      TabIndex        =   48
      Top             =   9840
      Width           =   735
   End
   Begin VB.Label numero2 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11640
      TabIndex        =   47
      Top             =   9840
      Width           =   1215
   End
   Begin VB.Label serie2 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10920
      TabIndex        =   46
      Top             =   9840
      Width           =   735
   End
   Begin VB.Label numero1 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   45
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label serie1 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   44
      Top             =   9840
      Width           =   735
   End
   Begin VB.Label tipo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8520
      TabIndex        =   43
      Top             =   240
      Width           =   495
   End
   Begin VB.Label flagcruce 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14880
      TabIndex        =   42
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label tipoclie 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14880
      TabIndex        =   41
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label flage 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14880
      TabIndex        =   40
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label txsubtotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   7320
      TabIndex        =   39
      Top             =   8640
      Width           =   1455
   End
   Begin VB.Label txdescuento 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   7320
      TabIndex        =   38
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label tximpuesto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   13080
      TabIndex        =   37
      Top             =   8280
      Width           =   1695
   End
   Begin VB.Label txneto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   7320
      TabIndex        =   36
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label acu 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14880
      TabIndex        =   35
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ntcant 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1680
      TabIndex        =   34
      Top             =   9000
      Width           =   855
   End
   Begin VB.Label txtotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   13080
      TabIndex        =   33
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Label estado 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14400
      TabIndex        =   32
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label17 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   0
      TabIndex        =   31
      Top             =   7920
      Width           =   3135
   End
   Begin VB.Label presenta 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   0
      TabIndex        =   30
      Top             =   8280
      Width           =   3135
   End
   Begin VB.Label Label15 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Observa"
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
      Left            =   9240
      TabIndex        =   29
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label14 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LocalDestino"
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
      Left            =   11640
      TabIndex        =   28
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Almacen Actual"
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
      Left            =   9240
      TabIndex        =   27
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nro.Dias"
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
      Left            =   6480
      TabIndex        =   26
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Left            =   6480
      TabIndex        =   25
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FormaPago"
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
      Left            =   3480
      TabIndex        =   24
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Transport."
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
      Left            =   6480
      TabIndex        =   23
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vendedor"
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
      Left            =   6480
      TabIndex        =   22
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3480
      TabIndex        =   21
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F.Entrega"
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
      Left            =   3480
      TabIndex        =   20
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F.Emision"
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
      Left            =   3480
      TabIndex        =   19
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo"
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
      Left            =   0
      TabIndex        =   18
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label tipo 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local Act"
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
      Left            =   9240
      TabIndex        =   17
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Numero"
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
      Left            =   0
      TabIndex        =   16
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label labelTIPONCD 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo Nota:"
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
      Left            =   9240
      TabIndex        =   235
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Menu dnu834 
      Caption         =   "&Nuevo"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu grba1 
      Caption         =   "&Grabar"
   End
   Begin VB.Menu gj88555 
      Caption         =   "&AgruparSaldo"
   End
   Begin VB.Menu sal8843 
      Caption         =   "&OrionV4.Saldos"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu dkie833 
      Caption         =   "&ProductosParaPedido"
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tfactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim bk2       As Variant

Dim xproducto As String

Dim opcion7   As Integer

Private Type campo_precio

    unidad As String
    factor As String
    precio As String
    costo As String
    margen As String
    stock As String

End Type

Dim c1                As String

Dim c2                As String

Dim c3                As String

Dim c4                As String

Dim c5                As String

Dim c6                As String

Dim c7                As String

Dim c8                As String

Dim c9                As String

Dim mytablexi         As New ADODB.Recordset

Dim mytablexsi        As New ADODB.Recordset

Dim mytablepeso       As New ADODB.Recordset

'NOP='S'  FLAG SALDO INICIAL CABECERA   L1 FLAG SALDO INICIAL DETALLE
Dim campo_precios(12) As campo_precio

Private Sub bo712_Click()

End Sub

Private Sub bodega_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(bodega) = 0 Then
        consulta_almacen
        Exit Sub

    End If

    found = busca_bodega("" & local1, "" & bodega, 0)

    If found = 0 Then
        bodega = ""
        bodega.SetFocus
        Exit Sub

    End If

    If localf.Visible = True Then
        localf.SetFocus
        Exit Sub

    End If

    If bodegaf.Visible = True Then
        bodegaf.SetFocus
        Exit Sub

    End If

    observa.SetFocus

End Sub

Private Sub bodega_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        transporte.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_almacen

    End If

End Sub

Private Sub bodegaf_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If Len(Trim(localf)) = 0 Then
        bodegaf = ""
        localf.SetFocus
        Exit Sub

    End If

    If KeyAscii <> 13 Then Exit Sub
    If ttipo = "Z" Then
        If Len(Trim(localf)) = 0 Then
            bodegaf = ""
            localf.SetFocus
            Exit Sub

        End If

        If Len(bodegaf) = 0 Then
            bodegaf.SetFocus
            Exit Sub

        End If

        found = busca_bodega("" & localf, "" & bodegaf, 1)

        If found = 0 Then
            bodegaf = ""
            Exit Sub

        End If

    End If

    observa.SetFocus

End Sub

Private Sub bodegaf_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        'localf.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        If Len(localf) = 0 Then
            localf.SetFocus
            Exit Sub

        End If

        consulta_almacenf

    End If

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

    'Exit Sub
    If Len(buffer) = 0 Then
        buffer = "%"
        buffer.SetFocus
        Exit Sub

    End If

    Command1_Click

End Sub

Private Sub buffer_KeyUp(KeyCode As Integer, Shift As Integer)

    'If KeyCode <> 13 And KeyCode <> 27 Then
    '   ejecuta 0
    'End If
End Sub

Private Sub CAJA_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    turno.SetFocus

End Sub

Private Sub caja_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        Frame7.Visible = False
        moneda.SetFocus
        Exit Sub

    End If

End Sub

Private Sub cajero_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(cajero) > 0 Then
        found = busca_cajero()

        If found = 0 Then
            MsgBox "NO existe Cajero", 48, "Aviso"
            cajero = ""
            cajero.SetFocus
            Exit Sub

        End If
   
    End If

    fechasunat.SetFocus

End Sub

Private Sub cajero_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        turno.SetFocus
        Exit Sub

    End If

End Sub

Private Sub cmdAddEntry_Click()

    If Frame4.Visible = True Then Exit Sub
    If dbgrid3.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    inicializa
    habilita_numero 0
    habilita_cabeza 0
    habilita_detalle 0
    ttipo.SetFocus

End Sub

Private Sub cmdCancelar_Click()

End Sub

Private Sub cmdDelete_Click()

End Sub

Private Sub cmdExit_Click()
    dlo132_Click

End Sub

Private Sub cmdGrabar_Click()

End Sub

Private Sub cmdPrint_Click()

End Sub

Private Sub cmdHelp_Click()

End Sub

Private Sub cmdGuiaRemision_Click()
    FrmGuiaRemision.Show 1
End Sub

Private Sub cmdSave_Click()
    grba1_Click

End Sub

Private Sub cmdSort_Click()

End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(Trim(codigo)) = 0 Then
        consulta_codigo
        Exit Sub

    End If

    If Len(codigo) > 0 Then
        'consulta_codigo
        found = busca_codigo()

        If found = 0 Then Exit Sub

    End If

    If bodegaf.Visible = False Then
        localf = local1

    End If

    fecha.SetFocus

End Sub

Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_codigo

    End If

    If KeyCode = &H26 Then
        If Numero.Enabled = True Then
            Numero.SetFocus

        End If

        Exit Sub

    End If

    If KeyCode = &H76 Then  'f7
        If tipoclie <> "V" And tipoclie <> "C" And tipoclie <> "P" Then
            Exit Sub

        End If

        If tipoclie = "C" Then
            tnclie.DBPROV = "clientes"
            tnclie.Show 1
            codigo.SetFocus

        End If

        If tipoclie = "P" Then
            tnclie.DBPROV = "proveedo"
            tnclie.Show 1
            codigo.SetFocus

        End If
   
    End If

End Sub

Private Sub Combo4_Click()

    If Len(tproducto) > 0 And Frame5.Visible = True Then
        dbgrid4.refresh
        tproducto = xproducto
        carga_dbgrid4

    End If

End Sub

Private Sub Command1_Click()
    ejecuta 1

End Sub

Sub ejecuta(sw As Integer)

    Dim buf       As String

    Dim buf1      As String

    Dim buf2      As String

    Dim xbuf      As String

    Dim xbuf2     As String

    Dim sfound    As String

    Dim rconsulta As New ADODB.Recordset

    'MsgBox sginicio
    If Check1.Value = 1 Then
        opcion1 = "45"

    End If

    buf2 = ""

    If tipoclie = "P" Then
        buf2 = "PROVEEDO"

    End If

    If tipoclie = "C" Then
        buf2 = "CLIENTES"

    End If

    If tipoclie = "V" Then
        buf2 = "VENDEDOR"

    End If

    'MsgBox opcion1
    buf1 = ""

    If opcion1 = "30" Then
        If Len(buffer) = 0 Then
            buf = "select Tipo,Serie,Numero,Codigo,Nombre,Fecha,Moneda as M,Total,Estado as E,FechaSunat from " & cgusuario & " where local='" & local1 & "' and tipo='" & ttipo & "' order by fecha"
        Else
            buf = "select Tipo,Serie,Numero,Codigo,Nombre,Fecha,Moneda as M,Total,Estado as E,FechaSunat from " & cgusuario & " where local='" & local1 & "' and tipo='" & ttipo & "' and " & Combo1 & " like '" & buffer & "%' order by fecha "

        End If

    End If

    If opcion1 = "443" Or opcion1 = "444" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from tlocal"
        Else
            buf = "select Nombre,Codigo from tlocal where  " & Combo1 & " like '" & buffer & "%'"

        End If

    End If
  
    If opcion1 = "21" Then
        If Len(buffer) = 0 Then
            buf = "select Descripcio,Tipo from Tipo where  anticipo='S'"
        Else
            buf = "select Descripcio,Tipo from tipo where  anticipo='S'  and " & Combo1 & " like '" & buffer & "%'"

        End If

    End If
   
    If opcion1 = "22" Or opcion1 = "23" Or opcion1 = "24" Then

        'If Len(buffer) = 0 Then
        'buf = "select Tipo,Numero,Fecha,Total,Moneda as M from recibo where usado<>'S' and tipo='" & retipo1 & "' and codigo='" & codigo & "'"
        'Else
        'buf = "select Tipo,Numero,Fecha,Total,Moneda as M from recibo where usado<>'S' and tipo='" & retipo1 & "' and codigo='" & codigo & "' and " & Combo1 & " like '" & buffer & "%'"
        'End If
    End If

    If opcion1 = "1" Then
        xbuf = " tipodoc='" & acu & "'"
        xbuf2 = ""

        If acu = "V" Then
            xbuf = " (tipodoc='A' or tipodoc='B' or tipodoc='C' or tipodoc='G' or tipodoc='D' )"

        End If

        If acu = "C" Then
            xbuf = " (tipodoc='J' or tipodoc='K' or tipodoc='L' or tipodoc='M' or tipodoc='P')"

        End If
      
        If Len(buffer) = 0 Then
            buf = "select Descripcio,Tipo from Tipo where " & xbuf
        Else
            buf = "select Descripcio,Tipo from tipo where " & xbuf & " and " & Combo1 & " like '%" & buffer & "%'"

        End If

    End If

    If opcion1 = "2" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo,Codigo1 from  " & buf2
        Else
            buf = "select Nombre,Codigo,Codigo1 from " & buf2 & " where " & Combo1 & " like '%" & buffer & "%'"

        End If

    End If

    If opcion1 = "3" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from Vendedor "
        Else
            buf = "select Nombre,Codigo from Vendedor where " & Combo1 & " like '%" & buffer & "%'"

        End If

    End If

    If opcion1 = "4" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from Transpor "
        Else
            buf = "select Nombre,Codigo from Transpor where " & Combo1 & " like '%" & buffer & "%'"

        End If

    End If
  
    If opcion1 = "5" Then
        If Len(buffer) = 0 Then
            buf = "select Descripcio,Fpago from Fpago where moneda='" & moneda & "'"
        Else
            buf = "select Descripcio,Fpago from Fpago where " & Combo1 & " like '%" & buffer & "%' and moneda='" & moneda & "'"

        End If

    End If

    If opcion1 = "6" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from Bodega where local='" & local1 & "'"
        Else
            buf = "select Nombre,Codigo from Bodega where local='" & local1 & "' and " & Combo1 & " like '%" & buffer & "%'"

        End If

    End If

    If opcion1 = "7" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from Bodega where local='" & localf & "'"
        Else
            buf = "select Nombre,Codigo from Bodega where local='" & localf & "' and " & Combo1 & " like '%" & buffer & "%'"

        End If

    End If

    If opcion1 = "8" Or opcion1 = "888" Or opcion1 = "50" Then
        '----------------------------
        sfound = tipo_costo("" & ttipo)

        'MsgBox sfound
        If Trim(sfound) = "C" Or Trim(sfound) = "P" Or Len(Trim(sfound)) = 0 Then
            If Len(buffer) = 0 Then
                buf = "select Producto.Descripcio,Producto.producto,Producto.Marca,Producto.Unidad as Und1,Producto.Factor as F,producto.Costou as Precio,Producto.Monedav as M,Producto.Familia as Fam,Producto.Subfamilia as Subfam,Producto.barras,producto.Igv from producto  "
            Else
                buf = "select Producto.Descripcio,Producto.producto,Producto.Marca,Producto.Unidad as Und1,Producto.Factor as F,Producto.Costou as Precio,Producto.Monedav as M,Producto.Familia as Fam,Producto.Subfamilia as Subfam,Producto.barras,producto.Igv from producto  WHERE   "
                buf = buf & Combo1 & " like '%" & buffer & "%'"

            End If

        End If
       
        '----------------------------
        If Trim(sfound) = "V" Then
            If Len(buffer) = 0 Then
                buf = "select Producto.Descripcio,Producto.producto,Producto.Marca,precios.Unidad1 as Und1,precios.Factor1 as F,precios.pventa1 as Precio,Producto.Monedav as M,Producto.Familia as Fam,Producto.Subfamilia as Subfam,Producto.barras,producto.Igv from producto  left join precios on producto.producto=precios.producto  where precios.local='" & local1 & "'"
            Else
                buf = "select Producto.Descripcio,Producto.producto,Producto.Marca,precios.Unidad1 as Und1,precios.Factor1 as F,Precios.pVenta1 as Precio,Producto.Monedav as M,Producto.Familia as Fam,Producto.Subfamilia as Subfam,Producto.barras,producto.Igv from producto left join precios on producto.producto=precios.producto WHERE  precios.local='" & local1 & "'  and "
                buf = buf & Combo1 & " like '%" & buffer & "%'"

            End If

        End If
         
    End If

    '---------------------------
      
    '---------------------------
    If opcion1 = "45" Then  'son compras a proveedores
        If Len(buffer) = 0 Then
            buf = "select Producto.descripcio,Producto.producto,producto.marca,producto.unidad as Und1,producto.Factor as F,Producto.Costou as Precio,producto.monedac as M,producto.familia,producto.Subfamilia,codprov.codigo from codprov left join producto on codprov.producto=producto.producto where codprov.codigo='" & codigo & "'"
        Else
            buf = "select Producto.descripcio,Producto.producto,producto.marca,producto.unidad as Und1,producto.Factor as F,Producto.Costou as Precio,producto.monedac as M,producto.familia,producto.Subfamilia,codprov.codigo from codprov left join producto on codprov.producto=producto.producto where codprov.codigo='" & codigo & "' and  descripcio like '%" & buffer & "%'"

        End If

    End If

    If Combo2 <> "%" Then
        buf = buf & " and " & Combo2 & " like '%" & buffer1 & "'"

    End If

    'MsgBox buf
    'MsgBox opcion1
    'MsgBox acu

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
    If opcion1 = "NC" Then  'son compras a proveedores
        If Len(buffer) = 0 Then
            buf = "select * from TIPONCD where tipo='NC'"
        Else
            buf = "select* from TIPONCD  where tipo='NC' AND " & Combo1 & " like '%" & buffer & "%'"

        End If

    End If

    If opcion1 = "ND" Then  'son compras a proveedores
        If Len(buffer) = 0 Then
            buf = "select * from TIPONCD where tipo='ND'"
        Else
            buf = "select* from TIPONCD where tipo='ND' and " & Combo1 & " like '%" & buffer & "%'"

        End If

    End If

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open buf, cn, adOpenStatic, adLockOptimistic

    If rconsulta.RecordCount = 0 Then
        rconsulta.Close

        If sw_consulta = 0 Then
            dlo132_Click
            Exit Sub

        End If

        buffer.SelStart = Len(buffer.Text)
        buffer.SetFocus
        Exit Sub

    End If
   
    Set DBGrid1.DataSource = rconsulta
    'refresca_precios
    sw_consulta = 1
   
    If opcion1 = "444" Or opcion1 = "443" Or opcion1 = "21" Or opcion1 = "1" Or opcion1 = "2" Or opcion1 = "3" Or opcion1 = "4" Or opcion1 = "5" Or opcion1 = "6" Or opcion1 = "7" Then
        DBGrid1.columns(0).Width = 4000
        DBGrid1.columns(1).Width = 2000

    End If

    If opcion1 = "22" Or opcion1 = "23" Or opcion1 = "24" Then
        DBGrid1.columns(0).Width = 1000
        DBGrid1.columns(1).Width = 1500
        DBGrid1.columns(2).Width = 1500
        DBGrid1.columns(3).Width = 1500
        DBGrid1.columns(4).Width = 700

    End If
               
    If opcion1 = "8" Or opcion1 = "888" Or opcion1 = "50" Or opcion1 = "45" Then
        DBGrid1.columns(0).Width = 5000
        DBGrid1.columns(1).Width = 1300
        DBGrid1.columns(2).Width = 1000
        DBGrid1.columns(3).Width = 900
        DBGrid1.columns(4).Width = 500
        DBGrid1.columns(5).Width = 900
        DBGrid1.columns(6).Width = 500
        DBGrid1.columns(7).Width = 800
        DBGrid1.columns(8).Width = 800
        DBGrid1.columns(9).Width = 1700

        'dbGrid1.Columns(10).Width = 500
    End If
               
    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
    If opcion1 = "NC" Or opcion1 = "ND" Then
        DBGrid1.columns(0).Width = 1300
        DBGrid1.columns(1).Width = 5000
        DBGrid1.columns(2).Width = 1300

    End If

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
               
    refresca_precios

    If sw = 1 Then
        DBGrid1.SetFocus

    End If

End Sub

Private Sub Command10_Click()

    Dim found As Integer

    If Len(cajero) > 0 Then
        found = busca_cajero()

        If found = 0 Then
            cajero = ""

        End If

    End If

    If Servicio <> "A" And Servicio <> "D" And Servicio <> "C" Then
        Servicio = "A"

    End If

    dlo132_Click

End Sub

Private Sub Command11_Click()

    Dim found As Integer

    If Len(cajero) > 0 Then
        found = busca_cajero()

        If found = 0 Then
            MsgBox "NO existe Cajero", 48, "Aviso"
            cajero = ""
            cajero.SetFocus
            Exit Sub

        End If

    End If

    If Len(caja) > 0 Then
        found = busca_caja()

        If found = 0 Then
            MsgBox "Caja No existe", 48, "Aviso"
            caja.SetFocus
            Exit Sub

        End If

    End If

    If Len(turno) > 0 Then
        found = busca_turno()

        If found = 0 Then
            MsgBox "Turno No existe", 48, "Aviso"
            turno.SetFocus
            Exit Sub

        End If

    End If

    If Servicio <> "A" And Servicio <> "D" And Servicio <> "C" Then
        Servicio = "A"

    End If

    'dlo132_Click
    Frame7.Visible = False
    dias.SetFocus

End Sub

Private Sub Command12_Click()
    Frame8.Visible = False
    dbgrid2.SetFocus

End Sub

Private Sub Command13_Click()

    On Error GoTo cmd5665_err

    Dim mytablex As New ADODB.Recordset

    If Len(codigo) = 0 Then
        MsgBox "No existe Codigo ", 48, "Aviso"
        Exit Sub

    End If

    If Len(producto) = 0 Then
        MsgBox "No existe Producto ", 48, "Aviso"
        Exit Sub

    End If

    If Len(rcodigo) = 0 Then
        MsgBox "No existe Codigo Proveedor rcodigo ", 48, "Aviso"
        Exit Sub

    End If
  
    mytablex.Open "SELECT *  FROM proveedo where codigo='" & codigo & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        MsgBox "No existe Proveedor ", 48, "Aviso"
        Exit Sub

    End If

    mytablex.Close
    mytablex.Open "SELECT *  FROM producto where producto='" & producto & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        MsgBox "No existe Producto ", 48, "Aviso"
        Exit Sub

    End If

    mytablex.Close
   
    mytablex.Open "SELECT *  FROM codprov where codigo='" & codigo & "' and producto='" & producto & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.AddNew
        mytablex.Fields("codigo") = "" & codigo
        mytablex.Fields("producto") = "" & producto
        mytablex.Fields("codigop") = "" & rcodigo
        mytablex.Update
    Else
        mytablex.Fields("codigop") = "" & rcodigo
        mytablex.Update

    End If

    mytablex.Close
    Exit Sub
cmd5665_err:
    MsgBox "Aviso en codprov ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command14_Click()
    Set mytablepeso = Nothing
    'If mytablepeso.State = 1 Then
    '   mytablepeso.Close
    'End If
    Frame6.Visible = False
    Exit Sub
    'Frame10.Visible = True
    'Frame10.Caption = "NUEVO"
    'inicializa_servicio
    'seproducto.Enabled = True
    'seproducto.SetFocus

End Sub

Private Sub Command15_Click()
    Data2.Recordset.Edit
    Data2.Recordset.Fields("cantidad") = Val(nsdx5)
    Data2.Recordset.Update

    'On Error GoTo cmd56122_err
    'Data5.Recordset.Delete
    'Exit Sub
    'cmd56122_err:
    'Exit Sub
End Sub

Private Sub Command16_Click()

End Sub

Private Sub Command17_Click()

    'Frame9.Visible = False
End Sub

Private Sub Command18_Click()

End Sub

Private Sub Command19_Click()

    'Frame10.Visible = False
End Sub

Private Sub Command2_Click()

    Dim sdx As Double

    dbgrid2.columns("t1") = Val(t1)
    dbgrid2.columns("t2") = Val(t2)
    dbgrid2.columns("t3") = Val(t3)
    dbgrid2.columns("t4") = Val(t4)
    dbgrid2.columns("t5") = Val(t5)
    dbgrid2.columns("t6") = Val(t6)
    dbgrid2.columns("t7") = Val(t7)
    dbgrid2.columns("t8") = Val(t8)
    dbgrid2.columns("t9") = Val(t9)
    dbgrid2.columns("t10") = Val(t10)
    dbgrid2.columns("t11") = Val(t11)
    dbgrid2.columns("t12") = Val(t12)
    dbgrid2.columns("t13") = Val(t13)
    dbgrid2.columns("t14") = Val(t14)
    dbgrid2.columns("t15") = Val(t15)
    dbgrid2.columns("t16") = Val(t16)
    sdx = Val(t1) + Val(t2) + Val(t3) + Val(t4) + Val(t5) + Val(t6) + Val(t7) + Val(t8) + Val(t9) + Val(t10) + Val(t11) + Val(t12) + Val(t13) + Val(t14) + Val(t15) + Val(t16)
    dbgrid2.columns("cantidad") = sdx
    calcula_igv 0
    Command3_Click

End Sub

Private Sub Command3_Click()
    dlo132_Click

End Sub

Private Sub Command4_Click()

    Dim sdx As Double

    dbgrid2.columns("observa1") = "" & Observa1
    dbgrid2.columns("observa2") = "" & observa2
    dbgrid2.columns("observa3") = "" & observa3
    dbgrid2.columns("observa4") = "" & observa4
    'sdx = Val(l1) + Val(l2) + Val(l3) + Val(l4)
    'DBGrid2.Columns(3) = sdx
    calcula_igv 0
    Command5_Click

End Sub

Private Sub Command5_Click()
    dlo132_Click

End Sub

Private Sub Command6_Click()
    dlo132_Click

End Sub

Private Sub Command7_Click()

    'cargar_productos_seleccionados
    'Frame4.Visible = False
    'buffer_KeyPress 27
End Sub

Private Sub Command8_Click()
    Frame5.Visible = False
    dbgrid2.Col = 4
    dbgrid2.Row = dbgrid2.VisibleRows - 2

    'DBGrid2.Col = 3
    dbgrid2.SetFocus

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim found As Integer

    Dim buf   As String

    Dim xtemp As Variant

    If KeyCode = &H70 Then  'f1
        If Len(DBGrid1.columns(0)) > 0 Then
            If opcion1 = "20" Then
                consulta_detalles

            End If

            Exit Sub

        End If

    End If

    If KeyCode = &H71 Then  'f2   cargar productos x bloque
        If Len(DBGrid1.columns(0)) > 0 Then
            If opcion1 = "8" Then
                consulta_bloques

            End If

            Exit Sub

        End If

    End If

    opcion3 = ""

    If KeyCode = &H72 Then  'f3
        If Len(DBGrid1.columns(0)) > 0 Then
            If opcion1 = "8" Then
                opcion3 = "1"
                xproducto = "" & DBGrid1.columns(1)
                tproducto = xproducto
                Combo4.ListIndex = 0
                carga_dbgrid4
                Exit Sub

            End If

        End If

    End If

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "30" Then  'ANULACION
            serie = Trim(DBGrid1.columns(1))
            Numero = Trim(DBGrid1.columns(2))
            Frame1.Visible = False
            Frame1.Enabled = False
            Numero.SetFocus
            numero_KeyPress 13

        End If

        If opcion1 = "21" Then

            'retipo1 = Trim(dbGrid1.columns(1))
            'Frame1.Visible = False
            'Frame1.Enabled = False
            'retipo1.SetFocus
            'retipo1_KeyPress 13
        End If

        If opcion1 = "443" Then
            local1 = Trim(DBGrid1.columns(1))
            Frame1.Visible = False
            Frame1.Enabled = False
            local1.SetFocus
            local1_KeyPress 13

        End If

        If opcion1 = "444" Then
            localf = Trim(DBGrid1.columns(1))
            Frame1.Visible = False
            Frame1.Enabled = False
            localf.SetFocus
            localf_KeyPress 13

        End If

        If opcion1 = "22" Then

            'renumero1 = Trim(dbGrid1.columns(1))
            'retotal1 = Trim(dbGrid1.columns(3))
            'suma_retotal
            'Frame1.Visible = False
            'Frame1.Enabled = False
            'renumero1.SetFocus
            'renumero1_KeyPress 13
        End If

        If opcion1 = "23" Then

            'renumero2 = Trim(dbGrid1.columns(1))
            'retotal2 = Trim(dbGrid1.columns(3))
            'suma_retotal
            'Frame1.Visible = False
            'Frame1.Enabled = False
            'renumero2.SetFocus
            'renumero2_KeyPress 13
        End If

        If opcion1 = "24" Then

            'renumero3 = Trim(dbGrid1.columns(1))
            'retotal3 = Trim(dbGrid1.columns(3))
            'suma_retotal
            'Frame1.Visible = False
            'Frame1.Enabled = False
            'renumero3.SetFocus
            'renumero3_KeyPress 13
        End If

        If opcion1 = "1" Then
            ttipo = Trim(DBGrid1.columns(1))
            Frame1.Visible = False
            Frame1.Enabled = False
            ttipo.SetFocus
            ttipo_KeyPress 13

        End If

        If opcion1 = "2" Then
            codigo = Trim(DBGrid1.columns(1))
            Frame1.Visible = False
            Frame1.Enabled = False
            codigo.SetFocus
            codigo_KeyPress 13

        End If

        ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
        If opcion1 = "NC" Or opcion1 = "ND" Then
            TIPONCD = Trim(DBGrid1.columns(0))
            Frame1.Visible = False
            Frame1.Enabled = False
            observa.SetFocus

        End If

        ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

        If opcion1 = "3" Then
            vendedor = Trim(DBGrid1.columns(1))
            Frame1.Visible = False
            Frame1.Enabled = False
            vendedor.SetFocus
            vendedor_KeyPress 13

        End If

        If opcion1 = "4" Then
            transporte = Trim(DBGrid1.columns(1))
            Frame1.Visible = False
            Frame1.Enabled = False
            transporte.SetFocus
            transporte_KeyPress 13

        End If

        If opcion1 = "5" Then
            fpago = Trim(DBGrid1.columns(1))
            Frame1.Visible = False
            Frame1.Enabled = False
            fpago.SetFocus
            fpago_KeyPress 13

        End If

        If opcion1 = "6" Then
            bodega = Trim(DBGrid1.columns(1))
            Frame1.Visible = False
            Frame1.Enabled = False
            bodega.SetFocus
            bodega_KeyPress 13

        End If

        If opcion1 = "7" Then
            'If Trim(dbGrid1.columns(1)) = Trim(bodega) Then
            '   MsgBox "Almacen Diferente ", 48, "Aviso"
            '   Exit Sub
            'End If
            bodegaf = Trim(DBGrid1.columns(1))
            Frame1.Visible = False
            Frame1.Enabled = False
            bodegaf.SetFocus
            bodegaf_KeyPress 13

        End If

        If opcion1 = "50" Then
            producto = Trim(DBGrid1.columns(1))
            Frame1.Visible = False
            Frame1.Enabled = False
            producto.SetFocus
            producto_KeyPress 13

        End If

        If opcion1 = "888" Then
            'seproducto = Trim("" & dbgrid1.columns(1))
            'sedescripcio = Trim("" & dbgrid1.columns(0))
            'seunidad = Trim("" & dbgrid1.columns(3))
            'sefactor = Trim("" & dbgrid1.columns(4))
            'seprecio = Trim("" & dbgrid1.columns(5))
            'Frame1.Visible = False
            'Frame1.Enabled = False
            'secantidad.SetFocus
            Exit Sub

        End If

        If opcion1 = "8" Or opcion1 = "45" Then
   
            If Len(Trim("" & dbgrid2.columns("producto"))) = 0 And Len(Trim("" & DBGrid1.columns(1))) > 0 Then
                found = verifica_doble(Trim("" & DBGrid1.columns(1)))

                If found = 1 Then
                    MsgBox "Producto ya seleccionado", 48, "Aviso"
                    DBGrid1.SetFocus
                    Exit Sub

                End If

                'MsgBox ""
      
                xtemp = dbgrid2.Row
                'Data2.Refresh
                dbgrid2.refresh
                'solo_ir_ultimo
                dbgrid2.Enabled = True
                dbgrid2.SetFocus

                If xtemp = -1 Then
                    xtemp = 0

                End If

                'MsgBox "XXX-XXX"
                dbgrid2.Row = xtemp
                dbgrid2.Col = 0
                dbgrid2.columns("producto") = Trim("" & DBGrid1.columns(1))
                found = busca_producto(Trim("" & dbgrid2.columns("producto")), 0, 0)

                If found = 0 Then
                    Exit Sub

                End If

                Frame1.Visible = False
                Frame1.Enabled = False
                'sumar_detalle
                'DBGrid2.Row = DBGrid2.VisibleRows - 1
                dbgrid2.Col = 4
                dbgrid2.SetFocus
                Exit Sub

            End If

        End If

    End If

End Sub

Sub consulta_bloques()
    'Frame4.Visible = True
    'Combo3.SetFocus
    Exit Sub

End Sub

Sub suma_retotal()

    'Dim sdx As Double
    'sdx = Val(retotal1) + Val(retotal2) + Val(retotal3)
    'retotal = Format(sdx, "0.00")
    'adetotal = Format(Val(retotal), "0.00")
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

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)

    Dim buf  As String

    Dim buf2 As String

    Exit Sub

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

Private Sub dbgrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    refresca_precios

End Sub

Private Sub dbgrid13_AfterColUpdate(ByVal ColIndex As Integer)

    'MsgBox ColIndex
    Select Case ColIndex

        Case 0

            'dbgrid13.columns("producto") = Trim("" & DBGrid2.columns("producto"))
        Case 1

        Case 2
            dbgrid13.columns("pesoneto") = Val("" & dbgrid13.columns("pesobruto")) - Val("" & dbgrid13.columns("tara"))

        Case 3
            dbgrid13.columns("pesoneto") = Val("" & dbgrid13.columns("pesobruto")) - Val("" & dbgrid13.columns("tara"))

        Case 4

    End Select

    dbgrid13.columns("total") = Val("" & dbgrid13.columns("pesoneto")) * Val("" & dbgrid13.columns("precio"))

    'sql_controlpeso Trim("" & DBGrid2.columns("producto"))
End Sub

Private Sub dbgrid13_BeforeColUpdate(ByVal ColIndex As Integer, _
                                     OldValue As Variant, _
                                     Cancel As Integer)
    dbgrid13.columns("producto") = Trim("" & dbgrid2.columns("producto"))
    dbgrid13.columns("precio") = Val("" & dbgrid2.columns("precio"))

End Sub

Private Sub DBGrid2_AfterColEdit(ByVal ColIndex As Integer)

    Select Case ColIndex

        Case 0

        Case 3

    End Select

End Sub

Private Sub dbgrid2_AfterColUpdate(ByVal ColIndex As Integer)

    Dim found As Integer

    Dim sdx   As Double

    Select Case ColIndex

        Case 0
            'found = busca_producto("" & DBGrid2.Columns(0), 0)
            'If found = 0 Then
            '   MsgBox "No existe producto", 48, "Aviso"
            '   Exit Sub
            'End If
            sumar_detalle
            dbgrid2.Col = 4
            dbgrid2.Row = dbgrid2.VisibleRows - 2
            dbgrid2.SetFocus

        Case 1
            dbgrid2.Col = 0
            dbgrid2.Row = dbgrid2.VisibleRows - 2
            dbgrid2.SetFocus
       
        Case 4
            'sdx = Val("" & DBGrid2.Columns(3)) * Val("" & DBGrid2.Columns(5))
            'DBGrid2.Columns(9) = Val(Format(sdx, "0.00"))
            'DBGrid2.Columns(7) = Val(Format(sdx, "0.00"))
            'calcula_igv
            'ir_ultimo
            sumar_detalle
            dbgrid2.Col = 6
            dbgrid2.Row = dbgrid2.VisibleRows - 2
            dbgrid2.SetFocus

        Case 6
            'sdx = Val("" & DBGrid2.Columns(3)) * Val("" & DBGrid2.Columns(5))
            'DBGrid2.Columns(9) = Val(Format(sdx, "0.00"))
            'DBGrid2.Columns(7) = Val(Format(sdx, "0.00"))
            'calcula_igv
            sumar_detalle
            dbgrid2.Col = 7
            dbgrid2.Row = dbgrid2.VisibleRows - 2
            dbgrid2.SetFocus

        Case 7
            'sdx = Val("" & DBGrid2.Columns(3)) * Val("" & DBGrid2.Columns(5))
            'DBGrid2.Columns(9) = Val(Format(sdx, "0.00"))
            'DBGrid2.Columns(7) = Val(Format(sdx, "0.00"))
            'calcula_igv
            sumar_detalle
            dbgrid2.Col = 7
            dbgrid2.Row = dbgrid2.VisibleRows - 2
            dbgrid2.SetFocus

        Case 8
            'If Val("" & DBGrid2.Columns(3)) > 0 Then
            '   sdx = Val("" & DBGrid2.Columns(7)) / Val("" & DBGrid2.Columns(3))
            '   DBGrid2.Columns(5) = Val(Format(sdx, "0.00"))
            '   DBGrid2.Columns(9) = Val("" & DBGrid2.Columns(7))
            '   calcula_igv
            sumar_detalle
            dbgrid2.Col = 0
            dbgrid2.Row = dbgrid2.VisibleRows - 1
            dbgrid2.SetFocus

            'End If
            ''' kenyo 08/11/2017 Percepcion en compras
        Case 10
            sumar_detalle
            dbgrid2.Col = 0
            dbgrid2.Row = dbgrid2.VisibleRows - 1
            dbgrid2.SetFocus
            ''' kenyo 08/11/2017 Percepcion en compras
            
    End Select

End Sub

Private Sub dbgrid2_BeforeColEdit(ByVal ColIndex As Integer, _
                                  ByVal KeyAscii As Integer, _
                                  Cancel As Integer)

    Dim found As Integer

    Select Case ColIndex

        Case 10, 11, 12, 13

            If Len("" & dbgrid2.columns("producto")) = 0 Then  '
                Cancel = True
                Exit Sub

            End If

            Exit Sub

    End Select

    If ColIndex > 10 Then
        Cancel = True
        Exit Sub

    End If

    Select Case ColIndex

        Case 2, 3
            Cancel = True
            Exit Sub

        Case 0

            If Len(codigo) = 0 Then
                MsgBox "debe existir cliente", 48, "Aviso"
                Cancel = True
                codigo.SetFocus
                Exit Sub

            End If

            If Len("" & dbgrid2.columns("producto")) > 0 Then  'si ya existe no cambiar
                Cancel = True
                Exit Sub

            End If

        Case 1

            If Len("" & dbgrid2.columns("producto")) = 0 Then  '
                Cancel = True
                Exit Sub

            End If

            If Len("" & dbgrid2.columns("descripcio")) = 0 Then  '
                Cancel = True
                Exit Sub

            End If
     
        Case 4

            If Len("" & dbgrid2.columns("producto")) = 0 Then  '
                Cancel = True
                Exit Sub

            End If

            If Len("" & dbgrid2.columns("linea")) > 0 Then  'ojo no se puede poner si es talla
                Cancel = True
                Exit Sub

            End If

        Case 6, 8, 7, 9, 10

            If Len("" & dbgrid2.columns("producto")) = 0 Then  '
                Cancel = True
                Exit Sub

            End If
     
    End Select

End Sub

Private Sub dbgrid2_BeforeColUpdate(ByVal ColIndex As Integer, _
                                    OldValue As Variant, _
                                    Cancel As Integer)

    Dim found As Integer

    Dim sdx   As Double

    'if bandera=""
    Select Case ColIndex

        Case 0

            If Len(dbgrid2.columns("producto")) = 0 Then
                Cancel = True
                Exit Sub

            End If

            If Len(dbgrid2.columns("producto")) > 14 Then
                Cancel = True
                Exit Sub

            End If

            found = verifica_doble("" & dbgrid2.columns("producto"))

            If found = 1 Then
                Cancel = True
                MsgBox "Producto ya Seleccionado", 48, "Aviso"
                Exit Sub

            End If

            found = busca_producto("" & dbgrid2.columns("producto"), 0, 0)

            If found = 0 Then
                Cancel = True

                'MsgBox "No existe producto", 48, "Aviso"
                If Mid$("" & dbgrid2.columns("producto"), 1, 1) <> "!" Then    'si es codigo de proveedor
                    consulta_producto "" & dbgrid2.columns("producto")

                End If

                'DBGrid2.Columns = 3
                Exit Sub

            End If

        Case 1

            If Len(dbgrid2.columns("producto")) = 0 Then
                Cancel = True
                Exit Sub

            End If

            If Len(dbgrid2.columns("descripcio")) = 0 Then
                Cancel = True
                Exit Sub

            End If

        Case 4

            If Len(dbgrid2.columns("producto")) = 0 Then
                Cancel = True
                Exit Sub

            End If

            If Not IsNumeric("" & dbgrid2.columns("cantidad")) Then
                Cancel = True
                Exit Sub

            End If

            sdx = Val("" & dbgrid2.columns("cantidad")) * Val("" & dbgrid2.columns("precio"))
            dbgrid2.columns("total") = sdx
            calcula_igv 0

        Case 6

            If Len(dbgrid2.columns("producto")) = 0 Then
                Cancel = True
                Exit Sub

            End If

            If Not IsNumeric(dbgrid2.columns("precio")) Then
                Cancel = True
                Exit Sub

            End If
     
            sdx = Val("" & dbgrid2.columns("precio"))

            If sinigv.Value = 1 Then
                sdx = Val("" & dbgrid2.columns("precio")) + Val("" & dbgrid2.columns("precio")) * Val("" & dbgrid2.columns("igv")) / 100
                dbgrid2.columns("precio") = sdx

            End If
     
            sdx = Val("" & dbgrid2.columns("cantidad")) * Val("" & dbgrid2.columns("precio"))
            dbgrid2.columns("total") = sdx
            calcula_igv 0
     
        Case 7

            If Len(dbgrid2.columns("producto")) = 0 Then
                Cancel = True
                Exit Sub

            End If

            If Not IsNumeric(dbgrid2.columns("deslipo")) Then
                Cancel = True
                Exit Sub

            End If

            sdx = Val("" & dbgrid2.columns("cantidad")) * Val("" & dbgrid2.columns("precio"))
            dbgrid2.columns("total") = sdx
            calcula_igv 0

        Case 8

            If Len(dbgrid2.columns("producto")) = 0 Then
                Cancel = True
                Exit Sub

            End If

            If Not IsNumeric(dbgrid2.columns("total")) Then
                Cancel = True
                Exit Sub

            End If

            If Val("" & dbgrid2.columns("cantidad")) = 0 Then
                Cancel = True
                Exit Sub

            End If

            sdx = Val("" & dbgrid2.columns("total"))

            If sinigv.Value = 1 Then
                sdx = Val("" & dbgrid2.columns("total")) + Val("" & dbgrid2.columns("total")) * Val("" & dbgrid2.columns("igv")) / 100
                dbgrid2.columns("total") = sdx

            End If
     
            sdx = Val("" & dbgrid2.columns("total")) / Val("" & dbgrid2.columns("cantidad"))
            dbgrid2.columns("precio") = sdx 'Val(Format(sdx, "0.00000"))
            calcula_igv 0

            'Case 14
            '     If Len(dbgrid2.columns("producto")) = 0 Then
            '        Cancel = True
            '        Exit Sub
            '     End If
            '     If Not IsNumeric(dbgrid2.columns("neto")) Then
            '        Cancel = True
            '        Exit Sub
            '     End If
            '     If Val("" & dbgrid2.columns("cantidad")) = 0 Then
            '        Cancel = True
            '        Exit Sub
            '     End If
            '     calcula_sinigv
            'calcula_igv 1
        Case 9

            If Len(dbgrid2.columns("producto")) = 0 Then
                Cancel = True
                Exit Sub

            End If

            dbgrid2.columns("tdetra") = Val(Format(Val("" & dbgrid2.columns("total")) * Val("" & dbgrid2.columns("detraccion")) / 100, "0.00"))
            calcula_igv 0

        Case 10

            If Len(dbgrid2.columns("producto")) = 0 Then
                Cancel = True
                Exit Sub

            End If

            dbgrid2.columns("tpercepcio") = Val(Format(Val("" & dbgrid2.columns("total")) * Val("" & dbgrid2.columns("percepcion")) / 100, "0.00"))

            If Trim(menup.Label10) = "ARGENTINA" Then
                dbgrid2.columns("tpercepcio") = Val(Format(Val("" & dbgrid2.columns("subtotal")) * Val("" & dbgrid2.columns("percepcion")) / 100, "0.00"))

            End If

            calcula_igv 0

        Case 11, 12, 13

            If Len(dbgrid2.columns("producto")) = 0 Then
                Cancel = True
                Exit Sub

            End If

            calcula_igv 0

    End Select

End Sub

Private Sub DBGrid2_ColEdit(ByVal ColIndex As Integer)

    Dim sdx As Double

    Select Case ColIndex

        Case 0

        Case 3
            
    End Select

End Sub

Private Sub dbgrid2_KeyDown(KeyCode As Integer, Shift As Integer)
    ver_presenta

    If KeyCode = 13 Then
        KeyCode = 0

        Select Case dbgrid2.Col

            Case 0

                If Len(dbgrid2.columns("producto")) > 0 Then
                    dbgrid2.Col = 4
                    Exit Sub

                End If

            Case 4

                If Val(dbgrid2.columns("cantidad")) > 0 Then
                    dbgrid2.Col = 8
                    Exit Sub

                End If

            Case 5

                If Val(dbgrid2.columns("cantidad")) > 0 Then
                    dbgrid2.Col = 6
                    Exit Sub

                End If

            Case 6

                If Val(dbgrid2.columns("precio")) > 0 Then
                    dbgrid2.Col = 8
                    Exit Sub

                End If

            Case 7

                If Val(dbgrid2.columns("precio")) > 0 Then
                    dbgrid2.Col = 8
                    Exit Sub

                End If

            Case 8

                If Val(dbgrid2.columns("total")) > 0 Then
                    dbgrid2.Col = 0
                    dbgrid2.Row = dbgrid2.VisibleRows - 1
                    Exit Sub

                End If

        End Select

    End If

End Sub

Private Sub dbgrid2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Exit Sub

        If bandera <> "Modifica" Then
            habilita_numero 0

        End If

        habilita_cabeza 0
        habilita_detalle 0
        observa.SetFocus
        Exit Sub

    End If

End Sub

Private Sub DBGrid2_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim found     As Integer

    Dim kproducto As String

    On Error GoTo cmd34_err

    ver_presenta

    If KeyCode = &H78 Then  'f9 consulta servicio

        'If Len(dbgrid2.columns("producto")) > 0 Then
        '   sql_servicio "" & dbgrid2.columns("producto")
        '   Exit Sub
        'End If
    End If

    If KeyCode = &H72 Then  'f3
        If Len(dbgrid2.columns("producto")) > 0 Then
            ingreso_locales

        End If

    End If

    If KeyCode = &H77 Then  'f1
        If Len(codigo) = 0 Then
            MsgBox "debe existir cliente", 48, "Aviso"
            codigo.SetFocus
            Exit Sub

        End If

        If Len(dbgrid2.columns("producto")) > 0 And dbgrid2.Col = 2 Then
            If Val("" & dbgrid2.columns("precio")) <= 0 Or Val("" & dbgrid2.columns("cantidad")) = 0 Then
                MsgBox "Deben existir costos y Cantidades Ingresados ", 48, "Aviso"
                dbgrid2.SetFocus
                Exit Sub

            End If

            kproducto = "" & dbgrid2.columns("producto")
            found = ver_cambio_precios(kproducto)
            dbgrid2.SetFocus
            Exit Sub

        End If

    End If

    If KeyCode = &H76 Then  'f7
        If Len(Trim("" & dbgrid2.columns("producto"))) > 0 Then
            xprodet.producto = Trim("" & dbgrid2.columns("producto"))
        Else

            'xprodet.producto = "ZZZ"
            'xprodet.counter = "2"
        End If

        xprodet.Show 1
        dbgrid2.SetFocus

    End If

    'If bandera = "Ver" Then Exit Sub
    If KeyCode = &H70 Then  'f1
        If Len(dbgrid2.columns("producto")) > 0 And dbgrid2.Col = 2 Then
            xproducto = "" & dbgrid2.columns("producto")
            tproducto = xproducto
            Combo4.ListIndex = 0
            carga_dbgrid4
            Exit Sub

        End If

    End If

    If KeyCode = &H72 And acu <> "3" Then 'f3   crea el codigo interno de cada proveedor
        If acu <> "C" Then Exit Sub
        Frame8.Visible = True
        producto = ""
        rcodigo = ""
        producto.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then

    End If

    If KeyCode = &H75 Then  'f6
        menu_carga

    End If

    'If KeyCode = &H77 Then  'f8 INGRESO DE INSUMOS
    '   tprodup.Caption = "Tabla de productos Insumos"
    '   tprodup.insumo.Value = 1
    '   tprodup.Show 1
    'End If
    If KeyCode = &H2E Then  'borrar linea
        If dbgrid2.Row = -1 Then
            MsgBox "No hay ningún registro para eliminar", vbInformation
            Exit Sub

        End If

        If MsgBox("Se va a eliminar el registro : está seguro ", vbExclamation + vbYesNo, "Eliminar") = vbYes Then
            Data2.Recordset.Delete

            If Data2.Recordset.EOF = True And Data2.Recordset.BOF = True Then
                Exit Sub

            End If

            ir_ultimo
            Data2.refresh
            'DBGrid2.Refresh
            dbgrid2.Col = 0
            dbgrid2.Row = dbgrid2.VisibleRows - 1
            dbgrid2.SetFocus
            Exit Sub

        End If

    End If

    If KeyCode = &H70 Then  'f1
        If Len(ttipo) = 0 Then
            ttipo.SetFocus
            Exit Sub

        End If

        found = busca_tipo(0)  'pone el acu

        If found = 0 Then
            ttipo.SetFocus
            Exit Sub

        End If

        found = busca_tipo(1)  'pone el acu

        If found = 0 Then
            ttipo.SetFocus
            Exit Sub

        End If

        If Len(dbgrid2.columns("producto")) = 0 Then
            consulta_producto ""

        End If

    End If

    If KeyCode = &H71 Then  'f2
        If Len(dbgrid2.columns("producto")) > 0 And Len(dbgrid2.columns("linea")) > 0 Then
            ingreso_tallas "" & dbgrid2.columns("linea")

        End If

    End If

    'If KeyCode = &H2D Then  'insert
    'If KeyCode = &H28 Then  'flecha abajo
    If KeyCode = &H28 Then  'flecha abajo inserta una nueva
        Exit Sub

        If dbgrid2.Col = 0 Then
            ir_ultimo

            If Len(dbgrid2.columns("producto")) > 0 And Len(dbgrid2.columns("descripcio")) > 0 And Len(dbgrid2.columns("unidad")) > 0 And Len(dbgrid2.columns("cantidad")) > 0 And Len(dbgrid2.columns("factor")) > 0 And Len(dbgrid2.columns("precio")) > 0 Then

                'Data2.Recordset.AddNew
                'Data2.Recordset.Update
            End If

            dbgrid2.Col = 0
            dbgrid2.Row = dbgrid2.VisibleRows - 1
            dbgrid2.SetFocus

        End If

    End If

    Exit Sub
cmd34_err:
    Exit Sub

End Sub

Private Sub DBGrid4_DblClick()
    DBGrid4_KeyDown 13, 0

End Sub

Private Sub DBGrid4_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim sdx As Double

    Dim buf As String

    If KeyCode = 27 Then
        If opcion3 = "1" Then
            Frame5.Visible = False
            DBGrid1.SetFocus
            Exit Sub

        End If

        Command8_Click
        Exit Sub

    End If

    If KeyCode = 13 Then
        If bandera = "Ver" Then Exit Sub

        'MsgBox opcion1
        If opcion3 = "1" Then
            Frame5.Visible = False
            DBGrid1.SetFocus
            Exit Sub

        End If

        'If opcion1 = "8" Then
        'If Len("" & DBGrid4.Columns(0)) > 0 And Val("" & DBGrid4.Columns(1)) > 0 And Len("" & DBGrid4.Columns(2)) > 0 Then
        'Data2.Recordset.Edit
        'Data2.Recordset.Fields("unidad") = "" & DBGrid4.Columns(0)
        'Data2.Recordset.Fields("factor") = "" & DBGrid4.Columns(1)
        'Data2.Recordset.Fields("precio") = "" & DBGrid4.Columns(3)
        'Data2.Recordset.Update
        dbgrid2.columns("unidad") = "" & dbgrid4.columns(0)
        dbgrid2.columns("factor") = Val("" & dbgrid4.columns(1))
        dbgrid2.columns("precio") = Val("" & dbgrid4.columns(2))
        buf = tipo_costo("" & ttipo)

        Select Case buf

            Case "C"
                dbgrid2.columns("precio") = Val("" & dbgrid4.columns(3)) ' / Val("" & DBGrid4.Columns(1))

            Case "V"
                dbgrid2.columns("precio") = Val("" & dbgrid4.columns(2))

        End Select

        sdx = Val("" & dbgrid2.columns("cantidad")) * Val("" & dbgrid2.columns("precio")) '* Val("" & DBGrid2.Columns("factor"))
        dbgrid2.columns("total") = sdx
        'MsgBox ""
        'Data2.Refresh
        calcula_igv 0
        sumar_detalle
        'DBGrid2.Col = 4
        'DBGrid2.Row = DBGrid2.VisibleRows - 2
        'DBGrid2.SetFocus
        Command8_Click

        'End If
        'End If
    End If

End Sub

Private Sub DBGrid4_UnboundReadData(ByVal RowBuf As MSDBGrid.RowBuffer, _
                                    StartLocation As Variant, _
                                    ByVal ReadPriorRows As Boolean)

    Dim dR            As Integer

    Dim row_num       As Integer

    Dim R             As Integer

    Dim rows_returned As Integer

    If ReadPriorRows Then
        dR = -1
    Else
        dR = 1

    End If

    If IsNull(StartLocation) Then
        If ReadPriorRows Then
            row_num = RowBuf.RowCount - 1
            'row_num = 9
        Else
            row_num = 0

        End If

    Else
        row_num = CLng(StartLocation) + dR

    End If

    rows_returned = 0

    For R = 0 To RowBuf.RowCount - 1

        If row_num < 0 Or row_num > 9 Then Exit For
        RowBuf.Value(R, 0) = campo_precios(row_num).unidad
        RowBuf.Value(R, 1) = campo_precios(row_num).factor
        RowBuf.Value(R, 2) = campo_precios(row_num).precio
        RowBuf.Value(R, 3) = campo_precios(row_num).costo
        RowBuf.Value(R, 4) = campo_precios(row_num).margen
        RowBuf.Value(R, 5) = campo_precios(row_num).stock
        RowBuf.Bookmark(R) = row_num
        row_num = row_num + dR
        rows_returned = rows_returned + 1
    Next R

    RowBuf.RowCount = rows_returned

End Sub

Private Sub DBGrid56_Click()

End Sub

Private Sub destino_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    Command11_Click

End Sub

Private Sub destino_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        partida.SetFocus
        Exit Sub

    End If

End Sub

Private Sub dias_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Val(dias) = 0 Then
        dias = "1"

    End If

    paridad.SetFocus

End Sub

Private Sub dias_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        fpago.SetFocus
        Exit Sub

    End If

End Sub

Private Sub djwewui_Click()

End Sub

Private Sub dkie833_Click()
    consulta_producto_inventario

End Sub

Private Sub dlo132_Click()

    Dim found As Integer

    On Error GoTo cmd891_err

    'If Frame9.Visible = True Then
    'Frame9.Visible = False
    'Exit Sub
    'End If

    If Frame6.Visible = True Then
        Frame6.Visible = False

        If mytablepeso.State = 1 Then
            mytablepeso.Close

        End If

        Exit Sub

    End If

    If Frame10.Visible = True Then
        Frame10.Visible = False
        Exit Sub

    End If

    If Frame7.Visible = True Then
        Frame7.Visible = False
        fechae.SetFocus
        Exit Sub

    End If

    If Frame4.Visible = True Then
        Frame4.Visible = False
        DBGrid1.SetFocus
        Exit Sub

    End If

    If dbgrid3.Visible = True Then
        cerrar_dbgrid3
        Exit Sub

    End If

    If Frame3.Visible = True Then
        Frame3.Visible = False
        dbgrid2.SetFocus
        Exit Sub

    End If

    If Frame2.Visible = True Then
        Frame2.Visible = False
        dbgrid2.SetFocus
        Exit Sub

    End If

    If Frame1.Visible = True Then
   
        If opcion1 = "444" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            localf.SetFocus
            Exit Sub

        End If

        If opcion1 = "443" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            local1.SetFocus
            Exit Sub

        End If

        If opcion1 = "21" Then
            'Frame1.Visible = False
            'Frame1.Enabled = False
            'retipo1.SetFocus
            Exit Sub

        End If

        If opcion1 = "GASTO" Then
            'Frame1.Visible = False
            'Frame1.Enabled = False
            'gasto.SetFocus
            Exit Sub

        End If
   
        If opcion1 = "22" Then
            'Frame1.Visible = False
            'Frame1.Enabled = False
            'renumero1.SetFocus
            Exit Sub

        End If

        If opcion1 = "23" Then
            'Frame1.Visible = False
            'Frame1.Enabled = False
            'renumero2.SetFocus
            Exit Sub

        End If

        If opcion1 = "24" Then
            'Frame1.Visible = False
            'Frame1.Enabled = False
            'renumero3.SetFocus
            Exit Sub

        End If

        If opcion1 = "30" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            serie.SetFocus
            Exit Sub

        End If

        If opcion1 = "1" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            ttipo.SetFocus
            Exit Sub

        End If

        If opcion1 = "2" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            codigo.SetFocus
            Exit Sub

        End If
   
        If opcion1 = "3" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            vendedor.SetFocus
            Exit Sub

        End If
   
        If opcion1 = "4" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            transporte.SetFocus
            Exit Sub

        End If

        If opcion1 = "5" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            fpago.SetFocus
            Exit Sub

        End If

        If opcion1 = "6" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            bodega.SetFocus
            Exit Sub

        End If

        If opcion1 = "7" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            bodegaf.SetFocus
            Exit Sub

        End If

        If opcion1 = "888" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            'seproducto.SetFocus
            Exit Sub

        End If
   
        ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
        If opcion1 = "ND" Or opcion1 = "NC" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            TIPONCD.SetFocus
            Exit Sub

        End If

        ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

        If opcion1 = "8" Or opcion1 = "45" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            'DBGrid2.Bookmark = bk2
            dbgrid2.Enabled = True
            dbgrid2.SetFocus
            Exit Sub

        End If

        If opcion1 = "50" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            producto.SetFocus
            Exit Sub

        End If

        Exit Sub

    End If

    If bandera = "Nuevo" Or bandera = "Modifica" Then
        found = salir_sin_grabar()

        If found = 0 Then
            Exit Sub

        End If

    End If

    tfactura.Hide
    Unload tfactura
    Exit Sub
cmd891_err:
    MsgBox "Error al salir " & error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub dnu834_Click()

    If Frame4.Visible = True Then Exit Sub
    If dbgrid3.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    If Frame7.Visible = True Then Exit Sub
    cmdAddEntry_Click

End Sub

Private Sub EC_Button1_MouseUp(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

    Dim found    As Integer

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    If fechai = "ORIONV4" Then
        proceso_v4
        Exit Sub

    End If

    If valida_fecha("" & fechai) = 0 Then
        fechai = ""
        fechai.SetFocus
        Exit Sub

    End If

    If valida_fecha("" & fechaf) = 0 Then
        fechaf = ""
        fechaf.SetFocus
        Exit Sub

    End If

    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub
    copiar_almacen0
    cn.Execute ("delete from almacen0 ")

    found = kardexactualizasi("" & local1, "%", "" & bodega, "" & fechai, "" & fechaf)

    If mytablexsi.State = 1 Then
        mytablexsi.Close

    End If

    mytablexsi.Open "SELECT * FROM almacen0", cn, adOpenKeyset, adLockOptimistic
    Set DataGrid1.DataSource = mytablexsi
    sdx = mytablexsi.RecordCount
    Label41 = "" & sdx

End Sub

Private Sub EC_Button2_MouseUp(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

    Dim found As Integer

    Dim sdx   As Double

    Dim sdx1  As Double

    Dim vr

    On Error GoTo cmd9099_err

    mytablexsi.Requery
    sdx1 = 0
    sdx = mytablexsi.RecordCount
    Label41 = "" & sdx
    Do

        If mytablexsi.EOF Then Exit Do
        If Val("" & mytablexsi("saldo")) > 0 Then
            found = busca_productosi(Trim("" & mytablexsi("producto")), 0, Val("" & mytablexsi("saldo")))

        End If

        vr = DoEvents()
        sdx1 = sdx1 + 1
        Label41 = "" & sdx & "-" & sdx1
        mytablexsi.MoveNext
    Loop
    Exit Sub
cmd9099_err:
    MsgBox "No existen Datos ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fecha_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(fecha) = 0 Then
        fecha = Format(Now, "dd/mm/yyyy")

    End If

    If Len(fechasunat) = 0 Then
        fechasunat = fecha

    End If

    fechae.SetFocus

End Sub

Private Sub fecha_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        codigo.SetFocus
        Exit Sub

    End If

End Sub

Private Sub fechae_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(fechae) = 0 Then
        fechae = Format(Now, "dd/mm/yyyy")

    End If

    If Len(fechasunat) = 0 Then
        fechasunat = fecha

    End If

    moneda.SetFocus

End Sub

Private Sub fechae_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        fecha.SetFocus
        Exit Sub

    End If

End Sub

Private Sub fechasunat_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(fechasunat) = 0 Then
        fechasunat = Format(Now, "dd/mm/yyyy")

    End If

    partida.SetFocus

End Sub

Private Sub fechasunat_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        cajero.SetFocus
        Exit Sub

    End If

End Sub

Sub inicializa_servicio()

    'seproducto = ""
    'sedescripcio = ""
    'seunidad = ""
    'sefactor = ""
    'secantidad = ""
    'seprecio = ""
    'setotal = ""
End Sub

Private Sub Form_Activate()

    Dim found As Integer

    'DBGrid2.columns(9).name = dicigv
    'If Len(caja) = 0 Then
    '   caja = "00"
    'End If
    'local1 = glocal
    'sql_detalle
    'sumar_detalle

    If acu <> "Z" Then
        Label14.Visible = False
        Label38.Visible = False
        localf.Visible = False
        bodegaf.Visible = False
        localf.Enabled = False
        bodegaf.Enabled = False
   
    End If
 
    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
    If acu = "E" Or acu = "F" Then  ' SI ES NOTA DE CREDITO O DEBITO
        labelTIPONCD.Visible = True
        TIPONCD.Visible = True
        documentorelacionado.Visible = True

    End If

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    vendedor = gusuario

    If cargado = "S" Then Exit Sub

    'racu = acu
    If bandera = "Nuevo" Then
        inicializa
        habilita_numero 0
        habilita_cabeza 0
        habilita_detalle 0
        ttipo.SetFocus

        If acu = "Q" Then
            ttipo = "Q"
            found = busca_tipo(0)  'pone el acu

            If found = 0 Then
                ttipo.SetFocus
                Exit Sub

            End If

            found = busca_tipo(6)

            If found = 0 Then
                serie.SetFocus
                Exit Sub

            End If

            fpago = "1"
            'sql_detalle
            'sumar_detalle
            Numero.SetFocus

        End If

        If acu = "H" Then
            ttipo = "H"
            found = busca_tipo(0)  'pone el acu

            If found = 0 Then
                ttipo.SetFocus
                Exit Sub

            End If

            found = busca_tipo(6)

            If found = 0 Then
                serie.SetFocus
                Exit Sub

            End If

            fpago = "1"
            'sql_detalle
            'sumar_detalle
            Numero.SetFocus

        End If

        If acu = "I" Then
            ttipo = "P"
            found = busca_tipo(0)  'pone el acu

            If found = 0 Then
                ttipo.SetFocus
                Exit Sub

            End If

            found = busca_tipo(6)

            If found = 0 Then
                serie.SetFocus
                Exit Sub

            End If

            fpago = "1"
            Numero.SetFocus

        End If
   
        If acu = "Z" Then
            'local1 = "01"
        
            ' 05/06/207 kenyo NO TIPO
            '  ttipo = "Z"
    
            found = busca_tipo(0)  'pone el acu

            If found = 0 Then
                ttipo.SetFocus
                Exit Sub

            End If

            found = busca_tipo(6)

            If found = 0 Then
                serie.SetFocus
                Exit Sub

            End If
      
            ' 05/06/207 kenyo NO CODIGO CLIENTE EN TRASLADO
            'codigo = "01"
      
            fpago = "1"
            Numero.SetFocus

        End If

    End If

    If bandera = "Ver" Then
        cmdSave.Enabled = False
        grba1.Enabled = False
        inicializa
        habilita_numero 0
        habilita_cabeza 0
        habilita_detalle 0
        local1 = zlocal
        ttipo = ztipo
        serie = zserie
        Numero = znumero
        found = busca_tipo(1)  'pone el acu
        found = busca_registro(1)

        If found = 0 Then
            MsgBox "No existe", 48, "Aviso"

        End If

        local1.Enabled = False
        ttipo.Enabled = False
        serie.Enabled = False
        Numero.Enabled = False
        sql_detalle
        sumar_detalle
        codigo.SetFocus
        dbgrid2.AllowUpdate = False

    End If

    If bandera = "Modifica" Then
        inicializa
        habilita_numero 0
        habilita_cabeza 0
        habilita_detalle 0
        local1 = zlocal
        ttipo = ztipo
        serie = zserie
        Numero = znumero
        found = busca_tipo(1)  'pone el acu
        found = busca_registro(1)

        If found = 0 Then
            MsgBox "No existe", 48, "Aviso"

        End If

        local1.Enabled = False
        ttipo.Enabled = False
        serie.Enabled = False
        Numero.Enabled = False
        sql_detalle
        sumar_detalle
        codigo.SetFocus
   
        ''''04/10/2017 kenyo Correcion duplicidad de traslados
        cn.Execute ("delete from detalle where local='" & localf & "' and tipo='TE' and serie='" & serie & "' and numero='" & Numero & "'")
        'cn.Execute ("delete from detalle where local='" & localf & "' and tipo='TS' and serie='" & serie & "' and numero='" & numero & "'")
        ''''04/10/2017 kenyo Correcion duplicidad de traslados
   
    End If

    cargado = "S"

End Sub

Private Sub Form_Load()

    'numcol = tempDBGrid.Columns.Count
    opcion7 = 0
    Combo3.AddItem "PROGRAMADO"
    Combo3.AddItem "TODOS"
    Combo3.AddItem "SALDOS<0"
    Combo3.AddItem "SALDOS<=0"
    Combo3.AddItem "SALDOS=0"
    Combo3.AddItem "SALDOS>0"
    Combo3.AddItem "SALDOS<=MINIMO"
    Combo3.AddItem "SALDOS<MAXIMO"
    Combo3.ListIndex = 0
    habilita_numero 1
    habilita_cabeza 1
    habilita_detalle 1
    carga_combo2

    

    'DBGrid2.Columns(0).Width = 2000
    'DBGrid2.Columns(1).Width = 3500
    'DBGrid2.Columns(2).Width = 800
    'DBGrid2.Columns(3).Width = 1000
    'DBGrid2.Columns(4).Width = 800
    'DBGrid2.Columns(5).Width = 1000
    'DBGrid2.Columns(6).Width = 800
    'DBGrid2.Columns(7).Width = 1000
    'DBGrid2.Columns(8).Width = 1000
    'DBGrid2.Columns(9).Width = 1000
               
    'DBGrid1.Columns(5).NumberFormat = "#,##0.00"
    'DBGrid2.Columns(3).NumberFormat = "###.##"
               
    'kenyo

    Frame1.Top = 60
    Frame1.Left = 0

    Frame6.Top = 705
    Frame6.Left = 30

    Frame10.Top = 705
    Frame10.Left = 30

    Frame5.Top = 765
    Frame5.Left = 4815

    Frame3.Top = 765
    Frame3.Left = 4815

    Frame7.Top = 765
    Frame7.Left = 4815

    Frame2.Top = 765
    Frame2.Left = 4815

    Frame8.Top = 765
    Frame8.Left = 4815

    Frame4.Top = 765
    Frame4.Left = 4815

End Sub

Sub inicializa()

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    On Error GoTo cmd3_err

    dflag.Value = 0
    txdetraccion = ""
    txisc = ""
    txestado = ""
    txivap = ""
    'almnombre = ""
    saldoini.Value = 0
    'gasto = ""
    'tipogasto = ""
    'seriegasto = ""
    'numerogasto = ""
    costopais = ""
    'agencia = ""
    'dua = ""
    hora = Format(Now, "hh:mm:ss")
    diferencia = ""
    costofactura = ""
    proveedorp = ""
    tproducto = ""
    precio = ""
    Servicio = "A"
    cajero = ""
    tflete = ""
    'xtotper = ""
    txpercepcio = ""
    'nbodega = ""
    fechasunat = ""
    opcion7 = 0

    Label17 = ""
    presenta = ""
    ttipo = ""
    serie = ""
    Numero = ""
    ntcant = ""
    txneto = ""
    txdescuento = ""
    txsubtotal = ""
    tximpuesto = ""
    txtotal = ""
    c1 = ""
    c2 = ""
    c3 = ""
    c4 = ""
    c5 = ""
    c6 = ""
    c7 = ""
    c8 = ""
    c9 = ""
    gravado = ""
    'adetotal = ""
    'acuenta = ""
    'retipo1 = ""
    'renumero1 = ""
    'renumero2 = ""
    'renumero3 = ""
    'retotal = ""
    'retotal1 = ""
    'retotal2 = ""
    'retotal3 = ""
    zona = ""
    Observa1 = ""
    observa2 = ""
    observa3 = ""
    observa4 = ""
    tipo1 = ""
    serie1 = ""
    serie2 = ""
    serie3 = ""
    serie4 = ""
    serie5 = ""
    serie6 = ""
    serie7 = ""

    numero1 = ""
    numero2 = ""
    numero3 = ""
    numero4 = ""
    numero5 = ""
    numero6 = ""
    numero7 = ""
    flagcruce = ""
    codigo = ""
    partida = ""
    destino = ""
    fecha = Format(Now, "dd/mm/yyyy")
    fechae = Format(Now, "dd/mm/yyyy")
    moneda = "S"
    vendedor = ""
    fpago = "1"
    transporte = ""
    dias = "1"
    bodega = ""
    localf = ""
    bodegaf = ""
    observa = ""
    estado = ""
    caja = ""

    If Len(Trim(bodega)) = 0 Then
        mytablex.Open "SELECT * FROM bodega where local='" & Trim(local1) & "'", cn, adOpenKeyset, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            bodega = Trim("" & mytablex.Fields("codigo"))

        End If

        mytablex.Close

    End If

    found = busca_bodega("" & local1, "" & bodega, 0)
    'local1 = glocal
    vendedor = gusuario

    paridad = "" & busca_paridadg(0)

    If Val(paridad) <= 0 Then
        paridad = "1"

    End If

    borrar_detalle_todo_registro
    'borra_importa
    sql_detalle
    Exit Sub
cmd3_err:
    MsgBox "Error en inicializa" & error$, 48, "Aviso"
    Exit Sub

End Sub

Function verificar_registro()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM " & cgusuario & " where  local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & Numero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        verificar_registro = 1

    End If

    mytablex.Close

End Function

Function busca_registro(sw As Integer)

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM " & cgusuario & " where  local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & Numero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        pone_registro mytablex
        busca_registro = 1

        If sw = 1 Then
            found = cargar_registrod()
            found = cargar_registrose()

            'found = cargar_importacion()
        End If

        If sw = 2 Then
            If "" & mytablex.Fields("yausado") <> "1" Then  'sino esta usado modificar
                If "" & mytablex.Fields("estado") = "2" Then
                    busca_registro = 2
                    found = cargar_registrod()
                    found = cargar_registrose()

                    'found = cargar_importacion()
                End If

                If "" & mytablex.Fields("estado") = "1" Then
                    busca_registro = 3

                End If

                If "" & mytablex.Fields("estado") = "0" Then
                    busca_registro = 4

                End If

            End If

        End If

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Sub pone_registro(mytablex As ADODB.Recordset)

    Dim found As Integer

    'Dim mytablex As New ADODB.Recordset
    'agencia = "" & mytablex.Fields("aduana")
    'dua = "" & mytablex.Fields("dua")
    hora = "" & mytablex.Fields("hora")
    tipoimp = "" & mytablex.Fields("tipoimp")
    serieimp = "" & mytablex.Fields("serieimp")
    numeroimp = "" & mytablex.Fields("numeroimp")

    caja = "" & mytablex.Fields("caja")
    turno = "" & mytablex.Fields("turno")
    Servicio = "" & mytablex.Fields("servicio")
    cajero = "" & mytablex.Fields("usuario")
    local1 = "" & mytablex.Fields("local")
    'adetotal = "" & mytablex.Fields("adetotal")
    'acuenta = "" & mytablex.Fields("acuenta")
    'retipo1 = "" & mytablex.Fields("retipo1")
    'renumero1 = "" & mytablex.Fields("renumero1")
    'renumero2 = "" & mytablex.Fields("renumero2")
    'renumero3 = "" & mytablex.Fields("renumero3")
    'retotal = "" & mytablex.Fields("retotal")
    'retotal1 = "" & mytablex.Fields("retotal1")
    'retotal2 = "" & mytablex.Fields("retotal2")
    'retotal3 = "" & mytablex.Fields("retotal3")
    '---
    zona = "" & mytablex.Fields("zona")
    ttipo = "" & mytablex.Fields("tipo")
    serie = "" & mytablex.Fields("serie")
    Numero = "" & mytablex.Fields("numero")
    codigo = "" & mytablex.Fields("codigo")
    partida = "" & mytablex.Fields("partida")
    destino = "" & mytablex.Fields("destino")
    fecha = "" & mytablex.Fields("fecha")
    fechasunat = "" & mytablex.Fields("fechasunat")
    fechae = "" & mytablex.Fields("fechae")
    moneda = "" & mytablex.Fields("moneda")
    vendedor = "" & mytablex.Fields("vendedor")
    Servicio = "" & mytablex.Fields("servicio")
    fpago = "" & mytablex.Fields("fpago")
    transporte = "" & mytablex.Fields("transporte")
    paridad = "" & mytablex.Fields("paridad")
    dias = "" & mytablex.Fields("dias")
    bodega = "" & mytablex.Fields("bodega")
    localf = "" & mytablex.Fields("localf")
    bodegaf = "" & mytablex.Fields("bodegaf")
    observa = "" & mytablex.Fields("observa")
    estado = "" & mytablex.Fields("estado")
    ntcant = "" & mytablex.Fields("nro_items")

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    If acu = "E" Or acu = "F" Then '  SI SE GENERA NC O ND
        TIPONCD = "" & mytablex.Fields("TIPONCD")
        estado_sunat = "" & mytablex.Fields("estado_sunat")
        tipo11 = "" & mytablex.Fields("tipo1")
        serie11 = "" & mytablex.Fields("serie1")
        numero11 = "" & mytablex.Fields("numero1")

    End If

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    tipo1 = "" & mytablex.Fields("tipo1")
    serie1 = "" & mytablex.Fields("serie1")
    serie2 = "" & mytablex.Fields("serie2")
    serie3 = "" & mytablex.Fields("serie3")
    serie4 = "" & mytablex.Fields("serie4")
    serie5 = "" & mytablex.Fields("serie5")
    serie6 = "" & mytablex.Fields("serie6")
    serie7 = "" & mytablex.Fields("serie7")

    numero1 = "" & mytablex.Fields("numero1")
    numero2 = "" & mytablex.Fields("numero2")
    numero3 = "" & mytablex.Fields("numero3")
    numero4 = "" & mytablex.Fields("numero4")
    numero5 = "" & mytablex.Fields("numero5")
    numero6 = "" & mytablex.Fields("numero6")
    numero7 = "" & mytablex.Fields("numero7")

    c1 = "" & mytablex.Fields("c1")
    c2 = "" & mytablex.Fields("c2")
    c3 = "" & mytablex.Fields("c3")
    c4 = "" & mytablex.Fields("c4")

    dflag.Value = 0

    If "" & mytablex.Fields("dflag") = "S" Then
        dflag.Value = 1

    End If

    saldoini.Value = 0

    If "" & mytablex.Fields("nop") = "S" Then
        saldoini.Value = 1

    End If

    'almnombre = ""
    'found = busca_codigo()
    'mytablex.Open "SELECT * FROM bodega where    codigo='" & bodega & "'", cn, adOpenKeyset, adLockOptimistic
    'If mytablex.RecordCount > 0 Then
    '   almnombre = "" & mytablex.Fields("nombre")
    'End If
    'mytablex.Close
    'nbodega = ""
    found = busca_bodega("" & local1, "" & bodega, 0)
    suma_retotal

End Sub

Sub grabando(mytablex As ADODB.Recordset)

    Dim buf As String

    On Error GoTo cmd781_err

    If dflag.Value = 1 Then
        mytablex.Fields("DFLAG") = "S"

    End If

    If importacion = "IMPORTACION" Then
        mytablex.Fields("tipoimp") = "I"

    End If

    If importacion = "COMERCIAL" Then
        mytablex.Fields("tipoimp") = "C"

    End If

    If importacion = "GASTOS" Then
        mytablex.Fields("tipoimp") = "G"

    End If

    mytablex.Fields("caja") = caja

    If Len(caja) = 0 Then
        mytablex.Fields("caja") = "00"

    End If

    If saldoini.Value = 1 Then
        mytablex.Fields("nop") = "S"
    Else
        mytablex.Fields("nop") = ""

    End If

    'mytablex.Fields("aduana") = Trim(agencia)
    'mytablex.Fields("dua") = Trim(dua)
    'mytablex.Fields("gasto") = Trim(cgasto)

    mytablex.Fields("turno") = turno
    mytablex.Fields("servicio") = Servicio
    'mytablex.Fields("adetotal") = Val(adetotal)
    'mytablex.Fields("acuenta") = Val(acuenta)

    'mytablex.Fields("retipo1") = retipo1
    'mytablex.Fields("renumero1") = renumero1
    'mytablex.Fields("renumero2") = renumero2
    'mytablex.Fields("renumero3") = renumero3
    'mytablex.Fields("retotal1") = Val(retotal1)
    'mytablex.Fields("retotal2") = Val(retotal2)
    'mytablex.Fields("retotal3") = Val(retotal3)
    'mytablex.Fields("retotal") = Val(retotal)
    mytablex.Fields("tflete") = Val(tflete)
    mytablex.Fields("zona") = zona
    mytablex.Fields("nombre") = Trim(Mid$("" & Label17, 1, 35))
    mytablex.Fields("estado") = "2"
    mytablex.Fields("yausado") = "0"
    mytablex.Fields("tipoclie") = tipoclie
    mytablex.Fields("tipo") = ttipo
    mytablex.Fields("serie") = serie
    mytablex.Fields("numero") = Numero
    mytablex.Fields("codigo") = codigo
    mytablex.Fields("partida") = partida
    mytablex.Fields("destino") = destino

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
    If acu = "E" Or acu = "F" Then
        mytablex.Fields("TIPONCD") = TIPONCD

        If estado_sunat = "" Then
            mytablex.Fields("estado_sunat") = 0

        End If

    End If

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    mytablex.Fields("nro_items") = Val(ntcant)

    If IsDate(fecha) Then
        mytablex.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
        mytablex.Fields("fechasunat") = Format(fecha, "dd/mm/yyyy")
        mytablex.Fields("fechae") = fechae
    Else
        mytablex.Fields("fecha") = Format(Now, "dd/mm/yyyy")
        mytablex.Fields("fechasunat") = Format(Now, "dd/mm/yyyy")
        mytablex.Fields("fechae") = Format(Now, "dd/mm/yyyy")

    End If

    mytablex.Fields("moneda") = moneda
    mytablex.Fields("vendedor") = vendedor

    mytablex.Fields("fpago") = fpago

    mytablex.Fields("transporte") = transporte
    mytablex.Fields("paridad") = Val(paridad)
    mytablex.Fields("dias") = Val(dias)
    mytablex.Fields("bodega") = bodega
    mytablex.Fields("localf") = localf
    mytablex.Fields("bodegaf") = bodegaf
    mytablex.Fields("observa") = observa
    mytablex.Fields("usuario") = "" & gusuario

    If Len(cajero) > 0 Then
        mytablex.Fields("usuario") = "" & cajero

    End If

    mytablex.Fields("acu") = "" & racu

    '''30/10/2017 Impresión de Tipo de documento en comprobantes
    'mytablex.Fields("documento") = "" & busca_tipocomprobante(ttipo)
    '''30/10/2017 Impresión de Tipo de documento en comprobantes

    '03/06/2017 KENYO
    'mytablex.Fields("acu1") = "" & acu1
    mytablex.Fields("acu1") = ""
    '03/06/2017 KENYO

    mytablex.Fields("flage") = "" & flage
    mytablex.Fields("hora") = Format(hora, "hh:MM:SS")
    mytablex.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")

    mytablex.Fields("fechasunat") = Format(fechasunat, "dd/mm/yyyy")
    mytablex.Fields("total") = Val("" & txtotal)
    'mytablex.Fields("recibe") = Val("" & txtotal)
    mytablex.Fields("descuento") = Val("" & txdescuento)
    mytablex.Fields("tisc") = Val("" & txisc)
    mytablex.Fields("tdetra") = Val("" & txdetraccion)
    mytablex.Fields("neto") = Val("" & txneto)
    mytablex.Fields("gravado") = Val("" & gravado)
    mytablex.Fields("impuesto") = Val("" & tximpuesto)
    mytablex.Fields("subtotal") = Val("" & txsubtotal)
    mytablex.Fields("percepcion") = Val("" & txpercepcio)

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
    If mytablex.Fields("acu") = "E" Or mytablex.Fields("acu") = "F" Then
        tipo1 = tipo11
        serie1 = serie11
        numero1 = numero11

    End If

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    mytablex.Fields("tipo1") = tipo1
    mytablex.Fields("serie1") = serie1
    mytablex.Fields("serie2") = serie2
    mytablex.Fields("serie3") = serie3
    mytablex.Fields("serie4") = serie4
    mytablex.Fields("serie5") = serie5
    mytablex.Fields("serie6") = serie6
    mytablex.Fields("serie7") = serie7

    mytablex.Fields("numero1") = numero1
    mytablex.Fields("numero2") = numero2
    mytablex.Fields("numero3") = numero3
    mytablex.Fields("numero4") = numero4
    mytablex.Fields("numero5") = numero5
    mytablex.Fields("numero6") = numero6
    mytablex.Fields("numero7") = numero7
    mytablex.Fields("local") = local1
    mytablex.Fields("c1") = Val(c1)
    mytablex.Fields("c2") = Val(c2)
    mytablex.Fields("c3") = Val(c3)
    mytablex.Fields("c4") = Val(c4)

    ''26/07/2017 kenyo CONDICION DE PAGO SISTEMA…
    If fpago = "3" Then
        mytablex.Fields("C9") = "3"
    Else
        mytablex.Fields("C9") = "1"

    End If

    ''26/07/2017 kenyo CONDICION DE PAGO SISTEMA…

    'si no es credito
    'grabar en acuenta igual a total
    buf = busca_fpagoc("" & fpago)  'credito ,letra

    If buf = "C" Or buf = "G" Then
    Else
        mytablex.Fields("acuenta") = Val("" & mytablex.Fields("total"))
        mytablex.Fields("adetotal") = 0

    End If

    'si es pedido grabar en acumulado clientes
    If "" & mytablex.Fields("acu") = "I" Then
        graba_acumulado_clientes "" & mytablex.Fields("codigo"), 1, Val("" & mytablex.Fields("total"))

    End If

    Exit Sub
cmd781_err:
    MsgBox "Aviso en grabando " + error$, 48, "Aviso"
    Exit Sub

End Sub

''30/10/2017 Impresión de Tipo de documento en comprobantes
Function busca_tipocomprobante(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT descripcio FROM tipo where  tipo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        busca_tipocomprobante = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close

End Function

''30/10/2017 Impresión de Tipo de documento en comprobantes

Sub graba_acumulado_clientes(buf As String, signo As Double, sumador As Double)

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM clientes where  codigo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        sdx = Val("" & mytablex.Fields("pedido")) + signo * sumador
        mytablex.Fields("pedido") = sdx
        mytablex.Update

    End If

    mytablex.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo cmd123_err

    'MsgBox "Hola"
    Exit Sub
cmd123_err:
    Exit Sub

End Sub

Private Sub formatode_Click()

End Sub

Private Sub fpago_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If moneda <> "S" And moneda <> "D" Then
        moneda = ""
        moneda.SetFocus
        Exit Sub

    End If

    If Len(fpago) = 0 Then
        consulta_fpago
        Exit Sub

    End If

    found = busca_fpago()

    If found = 0 Then
        fpago = ""
        Exit Sub

    End If

    Frame7.Visible = True
    caja.SetFocus

End Sub

Private Sub fpago_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        moneda.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_fpago

    End If

End Sub

Sub carga_combo2()

    Dim I As Integer

    Combo4.Clear

    For I = 1 To 9
        Combo4.AddItem Format(I, "00")
    Next I

    Combo4.ListIndex = 0

End Sub

Private Sub gj88555_Click()
    Frame10.Visible = True
    fechai = fecha
    fechaf = fecha
    Label41 = ""
    fechai.SetFocus

End Sub

Private Sub grba1_Click()

    Dim found    As Integer

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd78900_err

    If Frame7.Visible = True Then Exit Sub
    If Frame4.Visible = True Then Exit Sub
    If dbgrid3.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    dnu834.Enabled = False

    found = valida()

    If found = 0 Then
        MsgBox "Campos Invalidos", 48, "Aviso"
        dnu834.Enabled = True
        Exit Sub

    End If

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
    If acu = "E" Or acu = "F" Then   ' NOTA DE CREDITO
        If tipo11 = "" Or (tipo11 <> "1" And tipo11 <> "2") Then
            MsgBox "Verificar Tipo de documento relacionado", 48, "Aviso"
            tipo11.SetFocus
            Exit Sub

        End If
     
        If serie11 = "" Then
            MsgBox "Verificar Serie de documento relacionado", 48, "Aviso"
            serie11.SetFocus
            Exit Sub

        End If
     
        If numero11 = "" Then
            MsgBox "Verificar Número de documento relacionado", 48, "Aviso"
            numero11.SetFocus
            Exit Sub

        End If
      
        If Mid(serie, 1, 1) <> Mid(serie11, 1, 1) Then
            MsgBox "Serie de documento relacionado debe coincidir", 48, "Aviso"
            serie11.SetFocus
            Exit Sub

        End If
    
        If TIPONCD <> "01" And TIPONCD <> "02" And TIPONCD <> "03" And TIPONCD <> "04" And TIPONCD <> "05" And TIPONCD <> "06" And TIPONCD <> "07" And TIPONCD <> "08" And TIPONCD <> "09" Then          ' E nota de credito
            MsgBox "Tipo de nota No Existe", 48, "Aviso"
            TIPONCD.SetFocus
            Exit Sub

        End If
    
        If Len(observa) = 0 Then ' E nota de credito
            MsgBox "Ingrese observación", 48, "Aviso"
            observa.SetFocus
            Exit Sub

        End If
     
        Dim valor As Boolean

        Call Busca_comprobanteRelacionado_sunat(local1, serie11, numero11, tipo11, valor)
        
        If valor = False Then
            numero11.SetFocus
            Exit Sub

        End If

        ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
  
    End If

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    'MsgBox "grba1"
    'sumar_detalle
    'procesar_gastos
    'resumar_importacion
    If MsgBox("Desea Grabar", 1, "Aviso") <> 1 Then
        dbgrid2.SetFocus
        dnu834.Enabled = True
        Exit Sub

    End If

    'procesar_gastos

    If bandera = "Nuevo" Then  'adicionar
        If Len(Numero) = 0 Then
            mytablex.Open "SELECT * FROM tipo where    tipo='" & ttipo & "'", cn, adOpenKeyset, adLockOptimistic

            If mytablex.RecordCount > 0 Then  'si existe
                sdx = Val("" & mytablex.Fields("numero")) + 1
                Numero = "" & sdx

            End If

            mytablex.Close

        End If

akp:
        found = verificar_registro()

        If found = 1 Then
            sdx = Val(Numero) + 1
            Numero = "" & sdx
            GoTo akp

        End If

    End If

    If Not IsNumeric(Numero) Then
        Numero.SetFocus
        Exit Sub

    End If

    found = grabar()

    If found = 0 Then
        MsgBox "No se pudo grabar ", 48, "Aviso"
        dnu834.Enabled = True
        Exit Sub

    End If

    MsgBox "Proceso Grabado ", 48, "Aviso"

    If MsgBox("Desea Imprimir", 1, "Aviso") = 1 Then
        proceso_impresion1

    End If

    'habilita_numero 0
    'habilita_cabeza 0
    'habilita_detalle 0
    'inicializa   '0JO LE QUITE VERIFICA AUW NOPASE ESTO
    'If bandera = "Modifica" Then
    '   dlo132_Click
    'End If
    tfactura.Hide
    Unload tfactura
    'dlo132_Click
    Exit Sub
cmd78900_err:
    MsgBox "Ocurrio un aviso " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub hora_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    partida.SetFocus

End Sub

Private Sub Image1_Click()
    consulta_codigo
    Exit Sub

    If flagcruce = "S" Then
        tcrucedo.local1 = local1
        tcrucedo.tipo = tipo1
        tcrucedo.serie1 = serie1
        tcrucedo.serie2 = serie2
        tcrucedo.serie3 = serie3
        tcrucedo.serie4 = serie4
        tcrucedo.serie5 = serie5
        tcrucedo.serie6 = serie6
        tcrucedo.serie7 = serie7
        tcrucedo.numero1 = numero1
        tcrucedo.numero2 = numero2
        tcrucedo.numero3 = numero3
        tcrucedo.numero4 = numero4
        tcrucedo.numero5 = numero5
        tcrucedo.numero6 = numero6
        tcrucedo.numero7 = numero7
        tcrucedo.tipoclie = tipoclie
        tcrucedo.codigo = codigo
        tcrucedo.acu = racu
        tcrucedo.Show 1
    Else
        MsgBox "Tipo Documento sin permiso de Cruce", 48, "Aviso"

    End If

End Sub

Private Sub Image2_Click()
    'Dim found As Integer
    'If Len(gasto) = 0 Then
    '   gasto.SetFocus
    '   Exit Sub
    'End If
    'If Len(tipogasto) = 0 Then
    '   tipogasto.SetFocus
    '   Exit Sub
    'End If
    'If Len(seriegasto) = 0 Then
    '   seriegasto.SetFocus
    '   Exit Sub
    'End If
    'If Len(numerogasto) = 0 Then
    '   numerogasto.SetFocus
    '   Exit Sub
    'End If

    'found = busca_gasto()
    'If found = 0 Then
    '   MsgBox "Gasto no Existe", 48, "Aviso"
    '   gasto.SetFocus
    '   Exit Sub
    'End If
    'found = busca_facturagasto()
    'If found = 0 Then
    '   MsgBox "Factura Gasto no Existe", 48, "Aviso"
    '   tipogasto.SetFocus
    '   Exit Sub
    'End If
    'found = existe_facturagasto()
    'If found = 1 Then
    '   MsgBox "Factura Gasto Ya Existe", 48, "Aviso"
    '   tipogasto.SetFocus
    '   Exit Sub
    'End If
    'found = graba_facturagasto()

    'If found = 1 Then
    '   MsgBox "Factura Gasto Ya Existe", 48, "Aviso"
    '   tipogasto.SetFocus
    '   Exit Sub
    'End If
    'Label16_Click
End Sub

Private Sub Image4_Click()

    Dim ptipo    As String

    Dim pgasto   As String

    Dim pnumero  As String

    Dim pserie   As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd901200_err

    'pgasto = "" & mytablexi.Fields("gasto")
    'ptipo = "" & mytablexi.Fields("tipogasto")
    'pserie = "" & mytablexi.Fields("seriegasto")
    'pnumero = "" & mytablexi.Fields("numerogasto")
    'If mytablexi.State = 1 Then
    '   mytablexi.Close
    '   Set mytablexi = Nothing
    'End If
    'cn.Execute ("delete from gastofactura where  gasto='" & pgasto & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & numero & "' and tipogasto='" & ptipo & "' and seriegasto='" & pserie & "' and numerogasto='" & pnumero & "'")
    'Label16_Click
    Exit Sub
cmd901200_err:
    Exit Sub

End Sub

Private Sub Image6_Click()

    'Frame9.Visible = False
End Sub

Private Sub Label1_Click()
    cmdSort_Click

End Sub

Sub graba_servicio_tecnico()

    Dim I        As Integer

    Dim mytabley As New ADODB.Recordset

    Dim mytablex As Table

    mytabley.Open "SELECT * FROM serviciotecnico where  local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & Numero & "'", cn, adOpenKeyset, adLockOptimistic
    Set mytablex = mydbxglo.OpenTable(sgusuario)
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 1
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Fields("local") = local1
        mytabley.Fields("tipo") = ttipo
        mytabley.Fields("serie") = serie
        mytabley.Fields("numero") = Numero
         
        mytabley.Update
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close

End Sub

Function grabar()

    Dim rs               As Recordset

    Dim I                As Integer

    Dim pracu            As String

    Dim buf1             As String

    Dim found            As Integer

    Dim mytablex         As New ADODB.Recordset

    Dim mytabley         As New ADODB.Recordset

    Dim mytablez         As New ADODB.Recordset

    Dim mytablea         As New ADODB.Recordset

    Dim mytableb         As New ADODB.Recordset

    Dim mytablexy        As New ADODB.Recordset

    Dim te               As String

    Dim ts               As String

    Dim xc1              As Double

    Dim xc2              As Double

    Dim xc3              As Double

    Dim xc4              As Double

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
    Dim mysql            As String

    Dim my_nuevaCantidad As Double

    Dim k                As Integer

    Dim salida           As Boolean

    Dim my_codcliente    As String

    Dim my_ruc           As String

    Dim my_credito       As Boolean

    Dim my_idubigeo      As String

    Dim paso             As Boolean

    Dim file             As String

    Dim encontro         As Boolean

    Dim input_file       As String

    Dim my_CDR           As String

    Dim encontroSunat    As Boolean

    Dim my_codProducto   As String

    Dim my_costou        As String

    Dim my_cantidad      As Variant

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    Dim fila             As Integer

    Dim sw               As Integer

    Dim xbuf             As String

    On Error GoTo cmd761_err

    'graba cabecera
    If Not IsNumeric(Numero) Then
        Numero.SetFocus
        Exit Function

    End If

    sw = 0

    '03/06/2017 KENYO ORDEN DE COMPRA ACU1
    'acu1 = busca_tipox("" & tipo1)
    acu1 = ""
    '03/06/2017 KENYO ORDEN DE COMPRA ACU1

    If racu = "Z" Then  'abrir base datos traslado
        mytableb.Open "SELECT * FROM detalle where local='" & localf & "' and tipo='TS' and serie='" & serie & "' and numero='" & Numero & "'", cn, adOpenKeyset, adLockOptimistic

    End If

    'MsgBox dgusuariog
    'MsgBox cgusuario
    'MsgBox cgusuario
    xbuf = "SELECT * FROM " & cgusuario & " where  local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & Numero & "'"
   
    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open xbuf, cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then  'si existe
        mytablex.AddNew
        grabando mytablex
        mytablex.Update
        found = busca_tipo(7)   'graba  el numero
        graba_yausado_guia "0"
        'graba_gastofactura
        grabar = 1
    Else
        'mytablex.Edit
        grabando mytablex
        'MsgBox "" & mytablex.Fields("estado")
        mytablex.Update
        graba_yausado_guia "0"
        'graba_gastofactura
        grabar = 1

    End If

    mytablex.Close
    'MsgBox ""

    '-----grabar credito
    buf1 = busca_fpagoc("" & fpago)  'credito ,letra

    If buf1 = "C" Or buf1 = "G" Then
        If valida_flag("" & racu) = 1 Or valida_flag("" & racu) = 2 Then  'compras o ventas
      
            ''' 27/11/2017 Correcion duplicidad de guia salida venta en creditos
            If racu <> "T" Then
                grabar_cuentaxc

            End If

            ''' 27/11/2017 Correcion duplicidad de guia salida venta en creditos
   
        End If

    End If

    'MsgBox ""
    '----desapues ver si hubo adelantos
    'MsgBox ""
    'If Len(retipo1) > 0 Then
    '   If Len(renumero1) > 0 Then
    '   found = graba_adelantos("", "", retipo1, renumero1, "S")
    '   End If
    '   If Len(renumero2) > 0 Then
    '   found = graba_adelantos("", "", retipo1, renumero2, "S")
    '   End If
    '   If Len(renumero3) > 0 Then
    '   found = graba_adelantos("", "", retipo1, renumero3, "S")
    '   End If
    'End If
    '----si es letra hacer letra
    'buf1 = busca_fpagoc("" & fpago)
    'If buf1 = "G" Then
    '   If acu = "C" Or acu = "V" Or acu = "E" Or acu = "N" Or acu = "F" Or acu = "O" Then
    '      grabar_letras
    '   End If
    'End If
    '-----grabar forma de pago

    If valida_flag("" & racu) = 1 Or valida_flag("" & racu) = 2 Then  'compras o ventas
        found = graba_fpagov()

    End If

    '----------graba detalle------------------
    'borramos
    cn.Execute ("delete from " & dgusuariog & " where local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & Numero & "'")

    If racu = "3" Then 'si es servicio tecnico
        cn.Execute ("delete from serviciotecnico where local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & Numero & "'")
        graba_servicio_tecnico

    End If

    'ahora borramos en la base datos si es traslado
    If racu = "Z" Then
        cn.Execute ("delete from detalle where local='" & local1 & "' and tipo='TE' and serie='" & serie & "' and numero='" & Numero & "'")
        cn.Execute ("delete from detalle where local='" & localf & "' and tipo='TS' and serie='" & serie & "' and numero='" & Numero & "'")

    End If

    'MsgBox ""
    'GRABANDO EN detalle

    mytablexy.Open "SELECT * FROM " & dgusuariog & " where local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & Numero & "'", cn, adOpenKeyset, adLockOptimistic
    Data2.refresh
    Set rs = Data2.Recordset.Clone
    Do

        If rs.EOF Then Exit Do
        mytablexy.AddNew

        For I = 0 To rs.Fields.count - 1
            mytablexy.Fields(I) = rs.Fields(I)
        Next I

        'mytablexy.Fields("aduana") = Trim("" & agencia)
        'mytablexy.Fields("dua") = Trim("" & dua)

        If saldoini.Value = 1 Then
            mytablexy.Fields("l1") = "S"
        Else
            mytablexy.Fields("l1") = ""

        End If

        mytablexy.Fields("local") = "" & local1
        mytablexy.Fields("tipo") = "" & ttipo
        mytablexy.Fields("serie") = "" & serie
        mytablexy.Fields("numero") = "" & Numero
        mytablexy.Fields("vendedor") = "" & vendedor
        mytablexy.Fields("moneda") = "" & moneda
        mytablexy.Fields("bodega") = "" & bodega
        mytablexy.Fields("codigo") = "" & codigo
        mytablexy.Fields("localf") = "" & localf
        mytablexy.Fields("bodegaf") = "" & bodegaf
        mytablexy.Fields("acu") = "" & racu

        ''' 29/11/2017 Correción  General del Stock
        '''21/08/2017 kenyo Guia de Salida con Factura
        'mytablexy.Fields("acu1") = "" & acu1
        If mytablexy.Fields("acu1") = "A" Or mytablexy.Fields("acu1") = "B" Or mytablexy.Fields("acu1") = "C" Or mytablexy.Fields("acu1") = "D" Or mytablexy.Fields("acu1") = "G" Or mytablexy.Fields("acu1") = "S" Then
            mytablexy.Fields("acu1") = "S"
        Else
            mytablexy.Fields("acu1") = "" & acu1

        End If

        '''21/08/2017 kenyo Guia de Salida con Factura

        ''' 29/11/2017 Correción  General del Stock

        mytablexy.Fields("flage") = "" & flage
        mytablexy.Fields("tipoclie") = tipoclie
        mytablexy.Fields("usuario") = "" & gusuario

        If Len(cajero) > 0 Then
            mytablexy.Fields("usuario") = "" & cajero

        End If

        mytablexy.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
        'mytablexy.Fields("hora") = Format(Now, "hh:MM")

        mytablexy.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")

        '' 30/11/2017 Correción  General del Sistema Parte I. Error en reporte de comision

        mytablexy.Fields("comision") = "0"
        '' 30/11/2017 Correción  General del Sistema Parte I

        mytablexy.Fields("estado") = "2"
        mytablexy.Fields("caja") = caja

        If Len(caja) = 0 Then
            mytablexy.Fields("caja") = "00"

        End If

        mytablexy.Fields("turno") = turno
        mytablexy.Fields("servicio") = Servicio
        mytablexy.Update

        '----
        If racu = "Z" Then  'traslado
            mytableb.AddNew

            For I = 0 To rs.Fields.count - 1
                mytableb.Fields(I) = rs.Fields(I)
            Next I

            mytableb.Fields("local") = "" & local1
            mytableb.Fields("tipo") = "TS"
            mytableb.Fields("serie") = "" & serie
            mytableb.Fields("numero") = "" & Numero
            mytableb.Fields("vendedor") = "" & vendedor
            mytableb.Fields("moneda") = "" & moneda
            mytableb.Fields("bodega") = "" & bodega
            mytableb.Fields("localf") = "" & localf
            mytableb.Fields("bodegaf") = "" & bodegaf
            mytableb.Fields("acu") = "T"  'es salida
            mytableb.Fields("acu1") = "" & acu1
            mytableb.Fields("flage") = "" & flage
            mytableb.Fields("tipoclie") = tipoclie
            mytableb.Fields("codigo") = "" & codigo
            mytableb.Fields("usuario") = "" & gusuario

            If Len(cajero) > 0 Then
                mytableb.Fields("usuario") = "" & cajero

            End If

            mytableb.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
            'mytableb.Fields("hora") = Format(Now, "hh:MM")
            mytableb.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
            mytableb.Fields("estado") = "2"

            If Len(caja) = 0 Then
                mytableb.Fields("caja") = "00"

            End If

            mytableb.Update
            mytableb.AddNew
   
            For I = 0 To rs.Fields.count - 1
                mytableb.Fields(I) = rs.Fields(I)
            Next I
   
            mytableb.Fields("local") = localf '"" & codigo '& Mid$(codigo, 1, 3)
            mytableb.Fields("tipo") = "TE" '& ttipo
            mytableb.Fields("serie") = "" & serie
            mytableb.Fields("numero") = "" & Numero
   
            mytableb.Fields("vendedor") = "" & vendedor
            mytableb.Fields("moneda") = "" & moneda
            mytableb.Fields("bodega") = "" & bodegaf
            mytableb.Fields("localf") = "" & local1
            mytableb.Fields("bodegaf") = "" & bodega
            mytableb.Fields("acu") = "S"
            mytableb.Fields("acu1") = "" & acu1
   
            'para traslado no debe existir nada
            mytableb.Fields("flage") = "" & flage
            mytableb.Fields("tipoclie") = tipoclie
            mytableb.Fields("codigo") = "" & codigo
            mytableb.Fields("usuario") = "" & gusuario

            If Len(cajero) > 0 Then
                mytableb.Fields("usuario") = "" & cajero

            End If

            mytableb.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
            '  mytableb.Fields("hora") = Format(Now, "hh:MM")
            mytableb.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
            mytableb.Fields("estado") = "2"

            If Len(Trim(caja)) = 0 Then
                mytableb.Fields("caja") = "00"

            End If

            mytableb.Update
  
        End If

        'MsgBox ""
        If Len(codigo) > 0 Then
            If valida_flag("" & racu) = 2 Then  'compras
                found = crea_nuevos_proveedores("" & codigo, "" & rs.Fields("producto"), "" & rs.Fields("precio"), "" & fecha)

                'graba_costos rs, mytablea, mytabley, local1, bodega
                'MsgBox "costos"
                'descarga_saldo rs   'debe descaragr saldo
            End If

            If valida_flag("" & racu) = 1 Then  'ventas
                found = crea_nuevos_clientes("" & codigo, "" & rs.Fields("producto"), "" & rs.Fields("precio"), "" & fecha)

                'graba_costos rs, mytablea, mytabley, local1, bodega
                'MsgBox "costos"
                'descarga_saldo rs   'debe descaragr saldo
            End If

        End If

        grabar = 1
        '' 11/12/2017 SubReceta

        ' S ES PARA GUIA DE ENTRADA
        'If racu = "S" Or racu = "A" Or racu = "B" Or racu = "C" Or racu = "D" Or racu = "G" Then
        If racu = "A" Or racu = "B" Or racu = "C" Or racu = "D" Or racu = "G" Or racu = "T" Then

            '' 11/12/2017 Descontar stock de recetas desde Facturas de Ventas
            If mytablexy.Fields("acu1") <> "S" Then ' Si no es Guia que tiene factura no descarga nuevamente stock
                If verifica_receta("" & mytablexy.Fields("producto")) = 1 Then

                    Dim mytablezx As New ADODB.Recordset

                    '---------------------------------------
                    If mytablezx.State = 1 Then mytablezx.Close
                    mytablezx.Open "SELECT * FROM receta where producto='" & mytablexy.Fields("producto") & "'", cn, adOpenDynamic, adLockOptimistic

                    If mytablezx.RecordCount > 0 Then
                        Do

                            If mytablezx.EOF Then Exit Do
                     
                            ''' 11/12/2017 Descontar stock de recetas desde Facturas de Ventas
                            ''' 11/12/2017 SubReceta
                            If verifica_descargaproducto("" & mytablezx.Fields("productoi")) = 1 Then
                                ''' 11/12/2017 SubReceta
                     
                                mytablexy.AddNew
                       
                                mytablezx.Fields("descripcio") = "" & mytablezx.Fields("descripcio")
                                mytablexy.Fields("cantidad") = "" & mytablezx.Fields("cantidad")
                                mytablexy.Fields("producto") = "" & mytablezx.Fields("productoi")
                         
                                mytablexy.Fields("local") = local1
                                mytablexy.Fields("tipo") = "" & ttipo
                                mytablexy.Fields("serie") = "" & serie
                    
                                mytablexy.Fields("numero") = "" & Numero
                         
                                mytablexy.Fields("tipoclie") = "C"
                                mytablexy.Fields("moneda") = moneda
                                mytablexy.Fields("bodega") = bodega
                                mytablexy.Fields("bodegaf") = ""
                                mytablexy.Fields("acu") = acu
                                mytablexy.Fields("localf") = localf
               
                                mytablexy.Fields("flage") = ""
                                mytablexy.Fields("codigo") = "" & codigo
                                mytablexy.Fields("caja") = "" & caja
                                mytablexy.Fields("turno") = "" & turno
                                mytablexy.Fields("usuario") = "" & cajero
                                   
                                mytablexy.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
                                mytablexy.Fields("hora") = Format(Now, "hh:MM:ss")
                             
                                mytablexy.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
                                mytablexy.Fields("estado") = "2"
               
                                mytablexy.Fields("producto") = "" & mytablezx.Fields("productoi")
                                mytablexy.Fields("descripcio") = "" & mytablezx.Fields("descripcio")
                                mytablexy.Fields("unidad") = "" & mytablezx.Fields("unidad")
                                mytablexy.Fields("factor") = "" & mytablezx.Fields("factor")
                                mytablexy.Fields("cantidad") = Val("" & Data2.Recordset.Fields("cantidad")) * Val("" & mytablezx.Fields("cantidad")) * Val("" & mytablezx.Fields("factor"))
                                mytablexy.Fields("dua") = "R"  'flag que dice que es receta
                                mytablexy.Fields("acu") = "T"  'guia de salida
               
                                ''' kenyo 23/08/2017 Descontar stock de recetas en Recalculo de saldos
                                mytablexy.Fields("acu1") = ""  'guia de salida
                                ''' kenyo 23/08/2017 Descontar stock de recetas en Recalculo de saldos
               
                                mytablexy.Fields("precio") = 0
                                mytablexy.Fields("total") = 0
                                mytablexy.Fields("subtotal") = 0
                                mytablexy.Fields("impuesto") = 0
                                mytablexy.Fields("precio") = 0
                                mytablexy.Fields("neto") = 0
                                mytablexy.Fields("igv") = 0
                                mytablexy.Fields("descuento") = 0
                                mytablexy.Fields("subtotal") = 0
                                mytablexy.Fields("descuento") = 0
                         
                                mytablexy.Update
                     
                                ''' 11/12/2017 SubReceta
                            End If

                            ''' 11/12/2017 SubReceta
                            ''' 11/12/2017 Descontar stock de recetas desde Facturas de Ventas
                     
                            mytablezx.MoveNext
                        Loop

                    End If

                    mytablezx.Close

                    '---------------------------------------
                End If

            End If

            '' 11/12/2017 Descontar stock de recetas desde Facturas de Ventas
    
            '' 11/12/2017 SubReceta

        End If

        rs.MoveNext
    Loop

    found = valida_flag("" & racu)

    If found = 0 Then

    End If

    If found = 1 Or found = 2 Then
        'MsgBox "Hola"
        descarga_saldo "" & local1, "" & bodega, ttipo, serie, Numero, 0, ""

    End If

    If found = 3 Then
        'MsgBox ""
        descarga_saldo "" & local1, "" & bodega, "TS", serie, Numero, 0, "1"
        descarga_saldo "" & localf, "" & bodegaf, "TE", serie, Numero, 0, "1" 'mytablea productos

    End If

    'MsgBox ""

    If racu = "Z" Then
        mytableb.Close

    End If

    If estado_sunat = "" Then
        If acu = "E" Or acu = "F" Then
            Call busca_tipo_comprobante(local1, ttipo, serie, Numero, "", salida, ttipo, my_codcliente, my_acu)

            If salida = True Then
                'Call Datos_Empresa(my_struc_datos_empresa(), salida, 0)
                Call Datos_Empresa(my_struc_datos_empresa(), local1, salida, 0)
         
                my_ruc = my_struc_datos_empresa(0).codigo1
                Call b_ubigeo_receptor(my_codcliente, salida, my_struc_ubigeo_Receptor())
                Call b_ubigeo_emisor(my_ruc, salida, my_struc_ubigeo_Emisor())
                Call busca_cliente(my_codcliente, my_carga_busca_cliente(), salida, 0)
    
                ' Call b_tdescuento(serie1, numero1, my_acu, salida, my_struc_tventas())
       
                Call control_trasporte(local1, ttipo, serie, Numero, "", salida, my_struc_Etransporte(), 0)
    
                Call b_credito(serie, Numero, my_credito, my_struc_credito())
        
                'tIPO ES DOCUMENTO DE REFERENCIA
        
                If acu = "E" Then
                    Call estrae_nota_credito(my_ruc, local1, tipo1, serie, Numero, my_idubigeo, my_acu, my_struc_datos_empresa(), my_struc_ubigeo_Receptor(), my_carga_busca_cliente(), my_struc_credito(), paso, my_struc_ubigeo_Emisor(), file, ttipo)
                ElseIf acu = "F" Then
                    Call estrae_nota_debito(my_ruc, local1, tipo1, serie, Numero, my_idubigeo, my_acu, my_struc_datos_empresa(), my_struc_ubigeo_Receptor(), my_carga_busca_cliente(), my_struc_credito(), paso, my_struc_ubigeo_Emisor(), file, ttipo)

                End If

                Do While encontro = False
                    salida = FileExists("D:\ce_Input\FIRMADO\R_" & Left(file, (Len(file) - 10)) & ".txt")

                    If salida = True Then
                        Call busca_respuesta_electronica(file, "D:\ce_Input\FIRMADO\R_")
                        input_file = "D:\ce_Input\FIRMADO\R_" & Left(file, (Len(file) - 10)) & ".txt"
                        encontro = True

                    End If

                    salida = FileExists("D:\ce_Input\ERROR\R_" & Left(file, (Len(file) - 10)) & ".txt")

                    If salida = True Then
                        Call busca_respuesta_electronica(file, "D:\ce_Input\ERROR\R_")
                        input_file = "D:\ce_Input\ERROR\R_" & Left(file, (Len(file) - 10)) & ".txt"
                        encontro = True

                    End If

                Loop
    
                Call read_save_electronico(input_file, local1, serie, Numero, ttipo, my_acu, "")
        
                'Call busca_electronico(local1, serie1, numero1, _
                 tipo1, my_CDR)

                '    'Call graba_comprobanteE(input_file, local1, serie1, numero1, tipo1, file)
                '    Call Actualiza_cdr(input_file, local1, serie, numero, ttipo, acu, my_CDR)
                '    ttipo = tipo1
                '    serie = serie1
                '    numero = numero1
            End If

        End If

    End If

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    'mytablex.Close
    'mytablea.Close
    'mytablexy.Close
    'mytabley.Close
    Exit Function
cmd761_err:
    MsgBox "Aviso en grabar " + error$, 48, "Aviso"
    Exit Function

End Function

''' 11/12/2017 Descontar stock de recetas desde Facturas de Ventas
''' 11/12/2017 SubReceta
'Verifica si stock de producto se descarga del inventario
Function verifica_descargaproducto(buf As String)

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close

    '09/03/2018 Descontar stock de formularios
    'mytablex.Open "SELECT seccion FROM producto where seccion='INSUMO' and producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
    mytablex.Open "SELECT seccion FROM producto where producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
    '09/03/2018 Descontar stock de formularios

    If mytablex.RecordCount > 0 Then 'si existe
        verifica_descargaproducto = 1

    End If

    mytablex.Close

End Function

''' 11/12/2017 SubReceta
''' 11/12/2017 Descontar stock de recetas desde Facturas de Ventas

''''04/10/2017 kenyo Correcion duplicidad de traslados
Sub Actualiza_Traslados(buf1 As String, buf2 As String, buf3 As String)
    'Dim mytablex As New ADODB.Recordset
    '    If mytablefa.RecordCount > 0 Then
    '   mytablex.Open "delete FROM detalle where  tipo='" & buf1 & "' and  serie='" & buf2 & "' and numero='" & buf3 & "'", cn, adOpenKeyset, adLockOptimistic
    '    End If
    '
    'mytablex.Close

    Dim buf As String

    'PRODUCTO
    buf = "delete FROM detalle where  tipo='" & buf1 & "' and  serie='" & buf2 & "' and numero='" & buf3 & "'"
    cn.Execute (buf)

    ' actualiza_estadoti
End Sub

''''04/10/2017 kenyo Correcion duplicidad de traslados

Sub descarga_saldo(xlocal As String, _
                   xbodega As String, _
                   xtipo As String, _
                   xserie As String, _
                   xnumero As String, _
                   sw As Integer, _
                   tipoarchv As String)

    Dim sdx       As Double

    Dim signo     As Double

    Dim buf       As String

    Dim found     As Integer

    Dim sww       As Integer

    Dim mytablefa As New ADODB.Recordset

    Dim mytablex  As New ADODB.Recordset

    Dim mytabley  As New ADODB.Recordset

    On Error GoTo cmd19_err

    sww = 0
    'AHORA HAY QUE VALIDAR QUE no existe ya cruzado el documento----
    mytablefa.Open "SELECT * FROM " & cgusuario & " where  local='" & xlocal & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablefa.RecordCount > 0 Then  'si existe
        'If Len(mytablefa.Fields("tipo1")) > 0 And Len(mytablefa.Fields("serie1")) > 0 And Len(mytablefa.Fields("numero1")) > 0 Then
        '     found = ve_descarga("" & mytablefa.Fields("tipo1"))
        '     If found = 1 Then
        '      sww = 1
        '     End If
        'End If
   
        '25/06/2018 Testing Almacen General
        Dim xtipo1 As String

        xtipo1 = ""
        xtipo1 = mytablefa.Fields("tipo1")

        If Len(xtipo1) > 0 Then
            found = ve_descarga2(xtipo, xtipo1)

            If found = 1 Then
                sww = 1
            Else
                sww = 0

            End If

        End If

        '25/06/2018 Testing Almacen General

    End If

    mytablefa.Close

    buf = dgusuariog

    If tipoarchv = "1" Then
        buf = "detalle"

    End If

    mytablex.Open "SELECT * FROM " & buf & " where  local='" & xlocal & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then 'si existe
        Exit Sub

    End If

    'MsgBox ""
    'If permite_entrada_salida("" & mytablex.Fields("acu1")) = 1 Then 'si existe acu1 no descontar
    '   Exit Sub
    'End If
    Do

        If mytablex.EOF Then Exit Do
        '-------------------------------------------------
        signo = 1
      
        Select Case "" & mytablex.Fields("acu")
                '07/08/2018 Nota de credito final
                ' Case "S", "J", "K", "L", "M", "P", "E"
                ' signo = 1
            
                '07/08/2018 Nota de credito final
                          
                '25/06/2018 Testing Almacen General
                'Case "T", "A", "B", "C", "D", "G", "N"
                ' Case "T", "A", "B", "C", "D", "G", "N", "F"
             
                '07/08/2018 Nota de credito final
            Case "T", "A", "B", "C", "D", "G", "N", "F", "E"
             
                '07/08/2018 Nota de credito final
             
                '25/06/2018 Testing Almacen General
             
                signo = -1
             
                '19/06/2017 kenyo CORRECION STOCK ORDEN DE COMPRA
            Case "R" ' NO SUMA AL STOCL DE ORDEN DE COMPRA
                signo = 0

                '19/06/2017 kenyo CORRECION STOCK ORDEN DE COMPRA
        End Select

        'MsgBox signo
      
        If "" & mytablex.Fields("acu") = "P" Or "" & mytablex.Fields("acu") = "J" Or "" & mytablex.Fields("acu") = "K" Or "" & mytablex.Fields("acu") = "L" Or "" & mytablex.Fields("acu") = "M" Then 'compras varia el precios y costo
            graba_costos mytablex

        End If
      
        '-------------------------------------------------
        'busden:
        If sww = 0 Then
            If mytabley.State = 1 Then mytabley.Close
            mytabley.Open "select * from almacen where local='" & xlocal & "' and producto='" & Trim("" & mytablex.Fields("producto")) & "' and bodega='" & xbodega & "'", cn, adOpenDynamic, adLockOptimistic 'adOpenKeyset, adLockOptimistic

            'MsgBox mytabley.RecordCount
            If mytabley.RecordCount = 0 Then 'si existe
                'MsgBox ""
                mytabley.AddNew
                mytabley.Fields("local") = xlocal
                mytabley.Fields("producto") = "" & mytablex.Fields("producto")
                mytabley.Fields("bodega") = xbodega
                'mytabley.Fields("unidad") = "" & mytablex.Fields("unidad")
      
                sdx = Val("" & mytabley.Fields("saldo")) + signo * Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
    
                'MsgBox sdx
                mytabley.Fields("saldo") = sdx
                decarga_saldo_talla mytabley, mytablex, signo
                mytabley.Update
            Else

                If sw = 0 Then
                    'mytabley.Edit
                    'MsgBox ""

                    ''05/08/2017 kenyo Stock Inicial
                    'sdx = Val("" & mytabley.Fields("saldo")) + signo * Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
        
                    If explorap.Caption = "Documento Ingreso Saldo Inicial" Then
                        sdx = 0 + signo * Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
                    Else

                        If mytablex.Fields("ACU1") = "S" Then
                            sdx = Val("" & mytabley.Fields("saldo"))
                        Else
                            sdx = Val("" & mytabley.Fields("saldo")) + signo * Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))

                        End If

                    End If
   
                    '''21/08/2017 kenyo Guia de Salida con Factura

                    '''21/08/2017 kenyo Guia de Salida con Factura
   
                    ''05/08/2017 kenyo Stock Inicial
         
                    'MsgBox sdx
                    mytabley.Fields("saldo") = sdx
                    decarga_saldo_talla mytabley, mytablex, signo
                    mytabley.Update

                End If

                If sw = 1 Then
                    'mytabley.Edit
         
                    sdx = Val("" & mytabley.Fields("saldo")) - signo * Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
                    mytabley.Fields("saldo") = sdx
                    decarga_saldo_talla mytabley, mytablex, signo
                    mytabley.Update

                End If

                '-------------------------------
            End If

        End If 'fin sw sw

        '-------------------------------------------------
        mytablex.MoveNext
    Loop
    Exit Sub
cmd19_err:
    MsgBox "Aviso en descarga saldo " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub graba_costos(mytablex As ADODB.Recordset)

    Dim mytablexx As New ADODB.Recordset

    Dim mytabley  As New ADODB.Recordset

    Dim sdx3      As Double 'costo en una unidad del empaque

    Dim sdx4      As Double

    Dim sdx       As Double

    Dim coSmer    As Double

    Dim cossala   As Double

    Dim canstock  As Double

    Dim saldoant  As Double

    Dim asdx      As Double

    Dim bsdx      As Double

    On Error GoTo cmd23_err

    'MsgBox "L" & mytablex.Fields("local") & " P" & mytablex.Fields("producto") & " B" & mytablex.Fields("bodega")
    saldoant = 0
    mytablexx.Open "SELECT * FROM almacen where local='" & "" & mytablex.Fields("local") & "' and producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & "" & mytablex.Fields("bodega") & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablexx.RecordCount > 0 Then 'si existe
        saldoant = Val("" & mytablexx.Fields("saldo"))

    End If

    mytablexx.Close

    '''10/08/2017 kenyo Mejor Kardex Producto
    'sdx3 = (Val("" & mytablex.Fields("total")) / (Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))))   'costo empaque aque unidad
    
    ''''26/09/2017 kenyo Costos al por mayor Factura de Compra
    ' sdx3 = Val("" & mytablex.Fields("precio"))
    sdx3 = Val("" & mytablex.Fields("precio") / mytablex.Fields("factor"))
    ''''26/09/2017 kenyo Costos al por mayor Factura de Compra
    '''10/08/2017 kenyo Mejor Kardex Producto

    sdx4 = 0
    mytabley.Open "SELECT * FROM producto where  producto='" & "" & mytablex.Fields("producto") & "'", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount > 0 Then 'si existe

        'mytabley.Edit
        'costo ultimo poniendo
        If "" & mytablex.Fields("moneda") = "S" Then
            If "" & mytabley.Fields("monedac") = "S" Then
                sdx3 = sdx3

            End If

            If "" & mytabley.Fields("monedac") = "D" Then
                sdx3 = (sdx3 / Val(paridad))

            End If

        End If

        If "" & mytablex.Fields("moneda") = "D" Then
            If "" & mytabley.Fields("monedac") = "S" Then
                sdx3 = (sdx3 * Val(paridad))

            End If

            If "" & mytabley.Fields("monedac") = "D" Then
                sdx3 = sdx3

            End If

        End If

        'MsgBox sdx3
        'poniendo costo promedio
        'MsgBox sdx3
   
        If saldoant <= 0 Then
            sdx4 = sdx3

        End If

        If Val("" & mytabley.Fields("costop")) <= 0 Then
            sdx4 = sdx3

        End If
   
        '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
        '''10/08/2017 kenyo Mejor Kardex Producto
        'sdx3 = Val(Format(sdx3, "0.00"))
        sdx3 = Val(Format(sdx3, "0.0000"))
        '''10/08/2017 kenyo Mejor Kardex Producto
        '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
   
        coSmer = sdx3 * Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
        cossala = (Val("" & mytabley.Fields("costop"))) * saldoant
        canstock = saldoant + Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
   
        '13/08/2018 Integración FE - Pizzeria
        If sdx4 = 0 And canstock > 0 Then
            sdx4 = (coSmer + cossala) / canstock
      
            '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
            'sdx4 = Val(Format(sdx4, "0.00"))
            sdx4 = Val(Format(sdx4, "0.00000"))
            '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
      
            sdx4 = sdx4    '* Val("" & mytablex.Fields("factor1"))
            sdx3 = sdx3    '* Val("" & mytablex.Fields("factor1"))
        Else
            sdx4 = sdx4   '* Val("" & mytablex.Fields("factor1"))
            sdx3 = sdx3   '* Val("" & mytablex.Fields("factor1"))

        End If

        '13/08/2018 Integración FE - Pizzeria
   
        asdx = Val(Format(Val("" & mytabley.Fields("costou")), "0.00"))
   
        '''10/08/2017 kenyo Mejor Kardex Producto
        'bsdx = Val(Format(sdx3, "0.00"))
        bsdx = Val(Format(sdx3, "0.0000"))
        '''10/08/2017 kenyo Mejor Kardex Producto
   
        mytabley.Fields("ok") = ""

        If asdx <> bsdx Then
            If bsdx > 0 Then
                mytabley.Fields("ok") = "F"

            End If

        End If

        If sdx4 > 0 Then
            mytabley.Fields("costop") = sdx4

        End If

        If sdx3 > 0 Then
            'grabamos los costos anteiores
            mytabley.Fields("costoanterior2") = Val("" & mytabley.Fields("costoanterior1"))
            mytabley.Fields("costoanterior1") = Val("" & mytabley.Fields("costou"))
            mytabley.Fields("costou") = sdx3
      
            '13/08/2018 Integración FE - Pizzeria
            '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
            'actualiza_receta "" & mytablex.Fields("producto"), "" & sdx3
            If OpcionTipoCostoReceta() = "CP" Then
                actualiza_receta "" & mytablex.Fields("producto"), "" & sdx4
                actualiza_CostoTotalReceta "" & mytablex.Fields("producto"), "" & sdx4
            Else
                actualiza_receta "" & mytablex.Fields("producto"), "" & sdx3
                actualiza_CostoTotalReceta "" & mytablex.Fields("producto"), "" & sdx3

            End If

            '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
            '13/08/2018 Integración FE - Pizzeria
     
            If Val("" & mytabley.Fields("costoini")) = 0 Then
                mytabley.Fields("costoini") = sdx3

            End If

        End If

        If importacion = "IMPORTACION" Then
            mytabley.Fields("costopais") = Val("" & mytablex.Fields("canTdev")) / Val("" & mytablex.Fields("factor"))
            mytabley.Fields("costogasto") = Val("" & mytablex.Fields("flete"))

        End If

        actualizar_precios mytabley
        mytabley.Update
   
        '----Actualizar precio de venta si es margen automatico
   
    End If

    Exit Sub
cmd23_err:
    MsgBox "Aviso en Graba Costos " + error$, 48, "Aviso"
    Exit Sub

End Sub

Function valida()

    Dim found As Integer

    On Error GoTo cmd1934_err

    If Len(local1) = 0 Then
        MsgBox "Debe Existir Local1", 48, "Aviso"
        'local1.SetFocus
        Exit Function

    End If

    If Len(ttipo) = 0 Then
        ttipo.SetFocus
        Exit Function

    End If

    found = busca_tipo(0)  'valida el acu

    If found = 0 Then
        ttipo.SetFocus
        Exit Function

    End If

    'MsgBox "abc"
    If Len(serie) = 0 Then
        serie.SetFocus
        Exit Function

    End If

    If Len(Numero) > 0 Then
        If Numero.Enabled = True Then

            'If Not IsNumeric(numero) Then
            '   numero.SetFocus
            '   Exit Function
            'End If
        End If

    End If

    If bandera = "Nuevo" Then  'adicionar
        If Len(Numero) > 0 Then
            found = verificar_registro()

            If found = 1 Then
                MsgBox "Modo adicion,Ya existe el numero,cambie por otro", 48, "Aviso"
                Numero = ""
                Numero.SetFocus
                Exit Function

            End If

        End If

    End If

    If Len(codigo) > 0 Then
        found = busca_codigo()

        If found = 0 Then
            codigo.SetFocus
            Exit Function

        End If

    End If

    'MsgBox "abc"
    found = busca_tipo(3)   'valida el documento si obliga cruce

    If found = 8 Then
        If Len(serie1) = 0 Then
            MsgBox "Debe ingresar algun cruce de Documento", 48, "Aviso"
            codigo.SetFocus
            Exit Function

        End If

        If Len(numero1) = 0 Then
            MsgBox "Debe ingresar algun cruce de Documento", 48, "Aviso"
            codigo.SetFocus
            Exit Function

        End If

    End If

    If Len(vendedor) > 0 Then
        found = busca_vendedor()

        If found = 0 Then
            vendedor = ""
            vendedor.SetFocus
            Exit Function

        End If

    End If

    'MsgBox "abc"
    If Len(transporte) > 0 Then
        found = busca_transporte()

        If found = 0 Then
            transporte = ""
            transporte.SetFocus
            Exit Function

        End If

    End If

    If Len(fpago) = 0 Then
        fpago.SetFocus
        Exit Function

    End If

    found = busca_fpago()

    If found = 0 Then
        fpago = ""
        fpago.SetFocus
        Exit Function

    End If

    'MsgBox "abc"
    If Len(bodega) = 0 Then
        bodega.SetFocus
        Exit Function

    End If

    found = busca_bodega("" & local1, "" & bodega, 0)

    If found = 0 Then
        bodega = ""
        Exit Function

    End If

    'MsgBox "abc"
    If bodegaf.Enabled = True Then
        If Len(localf) = 0 Then
            localf = local1
            localf.SetFocus
            Exit Function

        End If

        found = busca_local1("" & localf)

        If found = 0 Then
            localf = local1
            localf.SetFocus
            Exit Function

        End If

        If Len(bodegaf) = 0 Then
            bodega = bodegaf
            bodegaf.SetFocus
            Exit Function

        End If

        If bodega = bodegaf Then

            'MsgBox "Almacenes no deben ser iguales ", 48, "Aviso"
            'bodegaf.SetFocus
            'Exit Function
        End If

        found = busca_bodega("" & localf, "" & bodegaf, 1)

        If found = 0 Then
            bodegaf = bodega
            bodegaf.SetFocus
            Exit Function

        End If

    End If

    'MsgBox "abc"
    If Len(fecha) <> 10 Then
        fecha = ""
        fecha.SetFocus
        Exit Function

    End If

    If Not IsDate(fecha) Then
        fecha = ""
        fecha.SetFocus
        Exit Function

    End If

    If Len(fechae) <> 10 Then
        fechae = ""
        fechae.SetFocus
        Exit Function

    End If

    If Not IsDate(fechae) Then
        fechae = ""
        fechae.SetFocus
        Exit Function

    End If

    If Val(paridad) <= 0 Then
        paridad = "1"

    End If

    If Len(fechasunat) = 0 Then
        fechasunat = Format(Now, "dd/mm/yyyy")

    End If

    If Not IsDate(fechasunat) Then
        fechasunat = Format(Now, "dd/mm/yyyy")

    End If

    If moneda <> "S" And moneda <> "D" Then
        moneda = "S"
        moneda.SetFocus
        Exit Function

    End If

    If Len(caja) > 0 Then
        found = busca_caja()

        If found = 0 Then
            MsgBox "Caja No existe", 48, "Aviso"
            Exit Function

        End If

    End If

    If Len(turno) > 0 Then
        found = busca_turno()

        If found = 0 Then
            MsgBox "Turno No existe", 48, "Aviso"
            Exit Function

        End If

    End If

    '' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
    'If ACU = "E" And Len(observa) = 0 Then ' E nota de credito
    '   MsgBox "Ingrese observación", 48, "Aviso"
    '   Exit Function
    'End If
    '
    'If ACU = "E" Then
    '    If TIPONCD <> "01" And TIPONCD <> "02" And TIPONCD <> "03" And TIPONCD <> "04" And TIPONCD <> "05" And TIPONCD <> "06" And TIPONCD <> "07" And TIPONCD <> "08" And TIPONCD <> "09" Then          ' E nota de credito
    '       MsgBox "Tipo de nota No Existe", 48, "Aviso"
    '       TIPONCD.SetFocus
    '       Exit Function
    '    End If
    'End If
    '
    '' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    If localf.Visible = False Then
        localf = local1

    End If

    If bodegaf.Visible = False Then
        bodegaf = bodega

    End If

    valida = 1

    Exit Function
cmd1934_err:
    MsgBox "Aviso en valida " + error$, 48, "Aviso"
    Exit Function

End Function

Private Sub modif2_Click()

End Sub

Private Sub Label16_Click()
    'resumar_importacion
    'Frame9.Visible = True
    'gasto.SetFocus

End Sub

Private Sub Label37_Click()

    'consulta_tipo
End Sub

Private Sub Label4_Click()

    Dim found As Integer

    found = leer_archivo_texto()
    fecha.SetFocus

End Sub

Private Sub Label45_Click()
    sql_controlpeso Trim("" & dbgrid2.columns("producto"))

End Sub

Private Sub Label49_Click()
    sumar_detalle

End Sub

Private Sub Label5_Click()

    Dim found As Integer

    found = guardar_fecha()
    fecha.SetFocus

End Sub

Private Sub local1_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

    If Len(local1) = 0 Then
        consulta_local1
        Exit Sub

    End If

    found = busca_local1("" & local1)

    If found = 0 Then
        'local1 = ""
        'loca11.SetFocus
        Exit Sub

    End If

    ttipo.SetFocus

End Sub

Private Sub local1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_local1

    End If

End Sub

Private Sub localf_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(localf) = 0 Then
        localf.SetFocus
        Exit Sub

    End If

    If Len(localf) > 0 Then
        found = busca_local1("" & localf)

        If found = 0 Then
            localf = ""
            localf.SetFocus
            Exit Sub

        End If

    End If

    bodegaf.SetFocus

End Sub

Private Sub localf_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_local2

    End If

    If KeyCode = &H26 Then
        bodega.SetFocus
        Exit Sub

    End If

End Sub

Private Sub moneda_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(moneda) = 0 Then
        moneda = "S"

    End If

    If moneda <> "S" And moneda <> "D" Then
        moneda.SetFocus
        Exit Sub

    End If

    fpago.SetFocus

End Sub

Private Sub moneda_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        fechae.SetFocus
        Exit Sub

    End If

End Sub

Private Sub numero_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If bandera = "Nuevo" Then
        If Len(Numero) = 0 Then

            'found = busca_tipo(9)
            'If found = 0 Then
            '   numero.SetFocus
            '   Exit Sub
            'End If
        End If

        If Len(Numero) > 0 Then
            found = verificar_registro()

            If found = 1 Then
                MsgBox "Modo adicion,Ya existe el numero,cambie por otro", 48, "Aviso"
                Numero = ""
                Numero.SetFocus
                Exit Sub

            End If

        End If

        codigo.SetFocus
        Exit Sub

    End If

    If Len(Numero) = 0 Then
        Numero.SetFocus
        Exit Sub

    End If

    codigo.SetFocus

End Sub

Private Sub numero_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        serie.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        If Len(ttipo) = 0 Then
            ttipo.SetFocus
            Exit Sub

        End If

    End If

End Sub

Private Sub observa_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    found = valida()

    If found = 0 Then
        MsgBox "Campos Invalidos ", 48, "Aviso"
        Exit Sub

    End If

    dbgrid2.Enabled = True
    sql_detalle
         
    dbgrid2.Row = dbgrid2.VisibleRows - 1
    dbgrid2.SetFocus

End Sub

Private Sub observa_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then

        'If dua.Enabled = True Then
        '   dua.SetFocus
        'End If
        'If dua.Enabled = False Then
        '   bodega.SetFocus
        'End If
    End If

End Sub

Private Sub observa1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    observa2.SetFocus

End Sub

Private Sub observa2_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    observa3.SetFocus

End Sub

Private Sub observa2_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        Observa1.SetFocus
        Exit Sub

    End If

End Sub

Private Sub observa3_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    observa4.SetFocus

End Sub

Private Sub observa3_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        observa2.SetFocus
        Exit Sub

    End If

End Sub

Private Sub observa4_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    Command4_Click

End Sub

Private Sub observa4_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        observa3.SetFocus
        Exit Sub

    End If

End Sub

Private Sub paridad_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    vendedor.SetFocus

End Sub

Private Sub paridad_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        dias.SetFocus
        Exit Sub

    End If

End Sub

Private Sub partida_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    destino.SetFocus

End Sub

Private Sub partida_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        If bodegaf.Enabled = True Then
            bodegaf.SetFocus
            Exit Sub

        End If

        bodega.SetFocus
        Exit Sub

    End If

End Sub

Private Sub producto_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    rcodigo.SetFocus

End Sub

Private Sub producto_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_rproducto

    End If

End Sub

Private Sub proveedorp_Change()

    Dim sdx As Double

    sdx = Val(Format(Val(costofactura), "0.00")) - Val(Format(Val("" & proveedorp), "0.00"))
    txestado = ""
    diferencia = Format(sdx, "0.00")

    If Val(costofactura) > Val(proveedorp) Then
        txestado = "SUBIO"

    End If

    If Val(costofactura) < Val(proveedorp) Then
        txestado = "BAJO"

    End If

End Sub

Private Sub rcodigo_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub sal8843_Click()
    Frame10.Visible = True
    fechai = "ORIONV4"
    fechaf = "ORIONV4"
    Label41 = ""
    fechai.SetFocus

End Sub

Private Sub serie_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(ttipo) = 0 Then
        ttipo.SetFocus
        Exit Sub

    End If

    found = busca_tipo(0)  'pone el acu

    If found = 0 Then
        ttipo.SetFocus
        Exit Sub

    End If

    If bandera = "Nuevo" Then
        found = busca_tipo(6)

        If found = 0 Then
            serie.SetFocus
            Exit Sub

        End If

    End If

    If Len(serie) = 0 Then
        MsgBox "Poner Numero serie ", 48, "Aviso"
        serie.SetFocus
        Exit Sub

    End If

    Numero.SetFocus

End Sub

Private Sub serie_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        ttipo.SetFocus
        Exit Sub

    End If

End Sub

Private Sub total_Click()
    sumar_detalle

End Sub

Private Sub servicio_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    Command11_Click
    dias.SetFocus

End Sub

Private Sub t1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t2.SetFocus

End Sub

Private Sub t10_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t11.SetFocus

End Sub

Private Sub t10_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        t9.SetFocus
        Exit Sub

    End If

End Sub

Private Sub t11_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t12.SetFocus

End Sub

Private Sub t11_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        t10.SetFocus
        Exit Sub

    End If

End Sub

Private Sub t12_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t13.SetFocus

End Sub

Private Sub t12_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        t11.SetFocus
        Exit Sub

    End If

End Sub

Private Sub t13_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t14.SetFocus

End Sub

Private Sub t13_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        t12.SetFocus
        Exit Sub

    End If

End Sub

Private Sub t14_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t15.SetFocus

End Sub

Private Sub t14_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        t13.SetFocus
        Exit Sub

    End If

End Sub

Private Sub t15_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t16.SetFocus

End Sub

Private Sub t15_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        t14.SetFocus
        Exit Sub

    End If

End Sub

Private Sub t16_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        t15.SetFocus
        Exit Sub

    End If

End Sub

Private Sub t2_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t3.SetFocus

End Sub

Private Sub t2_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        t1.SetFocus
        Exit Sub

    End If

End Sub

Private Sub t3_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t4.SetFocus

End Sub

Private Sub t3_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        t2.SetFocus
        Exit Sub

    End If

End Sub

Private Sub t4_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t5.SetFocus

End Sub

Private Sub t4_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        t3.SetFocus
        Exit Sub

    End If

End Sub

Private Sub t5_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t6.SetFocus

End Sub

Private Sub t5_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        t4.SetFocus
        Exit Sub

    End If

End Sub

Private Sub t6_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t7.SetFocus

End Sub

Private Sub t6_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        t5.SetFocus
        Exit Sub

    End If

End Sub

Private Sub t7_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t8.SetFocus

End Sub

Private Sub t7_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        t6.SetFocus
        Exit Sub

    End If

End Sub

Private Sub t8_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t9.SetFocus

End Sub

Private Sub t8_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        t7.SetFocus
        Exit Sub

    End If

End Sub

Private Sub t9_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t10.SetFocus

End Sub

Private Sub tl1_Click()

End Sub

Private Sub t9_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        t9.SetFocus
        Exit Sub

    End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    fechasunat.SetFocus

End Sub

' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
Private Sub TIPONCD_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

    If Len(TIPONCD) = 0 Then
        consulta_notacd
        Command1_Click
        Exit Sub

    End If

End Sub

' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
Private Sub TIPONCD_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_notacd
        Command1_Click

    End If

End Sub

' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

Private Sub transporte_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(transporte) > 0 Then
        found = busca_transporte()

        If found = 0 Then
            transporte = ""
            Exit Sub

        End If

    End If

    bodega.SetFocus

End Sub

Private Sub transporte_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        vendedor.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_transporte

    End If

    If KeyCode = &H76 Then  'f7
        ttranspo.Show 1

    End If

End Sub

Private Sub ttipo_Change()
If Trim(ttipo.Text) = "GR" Then
        cmdGuiaRemision.Visible = True
    Else
        cmdGuiaRemision.Visible = False

    End If
End Sub

Private Sub ttipo_KeyPress(KeyAscii As Integer)

    Dim found    As Integer

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    sdx = "" & sdx
    serie = ""
    Numero = ""
    tipo11 = ""
    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
 
    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

    If Len(ttipo) = 0 Then
        consulta_tipo
        Exit Sub

    End If

    found = busca_tipo(0)  'pone el acu

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
    If acu = "E" Or acu = "F" Then  ' SI ES NOTA DE CREDITO O DEBITO
        If Len(Numero) = 0 Then
            mytablex.Open "SELECT * FROM tipo where    tipo='" & ttipo & "'", cn, adOpenKeyset, adLockOptimistic

            If mytablex.RecordCount > 0 Then  'si existe
                sdx = Val("" & mytablex.Fields("numero")) + 1
                Numero = "" & sdx
                serie = "" & mytablex.Fields("serie")

            End If

            mytablex.Close

        End If
     
        If Mid(serie, 1, 1) = "B" Then
            tipo11 = "1"
        ElseIf Mid(serie, 1, 1) = "F" Then
            tipo11 = "2"
        Else
            tipo11 = ""

        End If
     
        codigo.SetFocus
    
        Exit Sub

    End If

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    If found = 0 Then Exit Sub
    'If tipoclie = "I" Then
    '   serie = "I01"
    'End If
    serie.SetFocus

End Sub

Private Sub ttipo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        'local1.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_tipo

    End If

End Sub

Private Sub turno_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    cajero.SetFocus

End Sub

Private Sub turno_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        caja.SetFocus
        Exit Sub

    End If

End Sub

Private Sub txestado_Click()

    On Error GoTo cmd9067_err

    If Len(Trim("" & dbgrid2.columns("producto"))) = 0 Then Exit Sub
    tproduct.codigo = Trim("" & dbgrid2.columns("producto"))
    tproduct.codigo.Enabled = False
    tproduct.ordename = "MODIFICA"
    tproduct.Show 1
    Exit Sub
cmd9067_err:
    MsgBox "Seleccione un dato", 48, "Aviso"
    Exit Sub

End Sub

Private Sub txtotal_Click()
    sumar_detalle

End Sub

Private Sub vendedor_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(vendedor) > 0 Then
        found = busca_vendedor()

        If found = 0 Then
            vendedor = ""
            Exit Sub

        End If

    End If

    transporte.SetFocus

End Sub

Private Sub vendedor_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        paridad.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_vendedor

    End If

    If KeyCode = &H76 Then  'f7
        tpersona.Show 1

    End If

End Sub

Sub consulta_tipo()

    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "Descripcio"
    Combo2.AddItem "Tipo"
    Combo2.ListIndex = 0
    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "Tipo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "1"
    sw_consulta = 0
    Command1_Click

End Sub

Sub consulta_gasto()

    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "Descripcio"
    Combo2.AddItem "gasto"
    Combo2.ListIndex = 0
    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "gasto"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "GASTO"
    sw_consulta = 0
    Command1_Click

End Sub

Sub consulta_codigo()
    cerrar_data1
    sw_consulta = 0
    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "Nombre"
    Combo2.AddItem "Codigo"
    Combo2.ListIndex = 0

    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.AddItem "Codigo1"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "2"

    If Len(Trim(buffer)) > 0 Then
        Command1_Click
        Exit Sub

    End If

    Set DBGrid1.DataSource = Nothing
    buffer.SetFocus

End Sub

Sub consulta_notacd()
    cerrar_data1
    sw_consulta = 0
    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "Nombre"
    Combo2.AddItem "Codigo"
    Combo2.ListIndex = 0

    Combo1.Clear
    Combo1.AddItem "Codigo"
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "tipo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus

    If acu = "E" Then
        opcion1 = "NC"
    ElseIf acu = "F" Then
        opcion1 = "ND"

    End If

    If Len(Trim(buffer)) > 0 Then
        Command1_Click
        Exit Sub

    End If

    Set DBGrid1.DataSource = Nothing
    buffer.SetFocus

End Sub

Sub consulta_agencia()
    cerrar_data1
    sw_consulta = 0
    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "Nombre"
    Combo2.AddItem "Codigo"
    Combo2.ListIndex = 0

    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.AddItem "Codigo1"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "AGENCIA"
    Command1_Click

End Sub

Sub consulta_local1()
    cerrar_data1
    sw_consulta = 0
    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "Nombre"
    Combo2.AddItem "Codigo"
    Combo2.ListIndex = 0
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "443"
    Command1_Click

End Sub

Sub consulta_local2()
    cerrar_data1
    sw_consulta = 0
    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "Nombre"
    Combo2.AddItem "Codigo"
    Combo2.ListIndex = 0
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "444"
    Command1_Click

End Sub

Sub consulta_vendedor()
    cerrar_data1
    sw_consulta = 0
    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "Nombre"
    Combo2.AddItem "Codigo"
    Combo2.ListIndex = 0
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "3"
    Command1_Click

End Sub

Sub consulta_retipo1()
    cerrar_data1
    sw_consulta = 0
    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "Tipo"
    Combo2.AddItem "Numero"
    Combo2.ListIndex = 0
    Combo1.Clear
    Combo1.AddItem "Tipo"
    Combo1.AddItem "Numero"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "21"
    Command1_Click

End Sub

Sub consulta_adelanto1()
    cerrar_data1
    sw_consulta = 0
    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "Tipo"
    Combo2.AddItem "Numero"
    Combo2.ListIndex = 0
    Combo1.Clear
    Combo1.AddItem "Tipo"
    Combo1.AddItem "Numero"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "22"
    Command1_Click

End Sub

Sub consulta_adelanto2()
    cerrar_data1
    sw_consulta = 0
    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "Tipo"
    Combo2.AddItem "Numero"
    Combo2.ListIndex = 0
    Combo1.Clear
    Combo1.AddItem "Tipo"
    Combo1.AddItem "Numero"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "23"
    Command1_Click

End Sub

Sub consulta_adelanto3()
    cerrar_data1
    sw_consulta = 0
    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "Tipo"
    Combo2.AddItem "Numero"
    Combo2.ListIndex = 0
    Combo1.Clear
    Combo1.AddItem "Tipo"
    Combo1.AddItem "Numero"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "24"
    Command1_Click

End Sub

Sub consulta_transporte()
    cerrar_data1
    sw_consulta = 0
    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "Nombre"
    Combo2.AddItem "Codigo"
    Combo2.ListIndex = 0

    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "4"
    Command1_Click

End Sub

Sub consulta_fpago()
    cerrar_data1
    sw_consulta = 0
    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "Descripcio"
    Combo2.AddItem "Fpago"
    Combo2.ListIndex = 0

    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "Fpago"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "5"
    Command1_Click

End Sub

Sub consulta_almacen()
    cerrar_data1
    sw_consulta = 0
    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "Nombre"
    Combo2.AddItem "Codigo"
    Combo2.ListIndex = 0

    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "6"
    Command1_Click

End Sub

Sub consulta_almacenf()
    cerrar_data1
    sw_consulta = 0
    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "Nombre"
    Combo2.AddItem "Codigo"
    Combo2.ListIndex = 0

    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "7"
    Command1_Click

End Sub

Function busca_tipo(sw As Integer)

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    Dim buf      As String

    'Label16 = ""
    racu = ""
    buf = ""
    mytablex.Open "SELECT * FROM tipo where   tipo='" & ttipo & "'" & buf, cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Function

    End If
   
    If acu = "V" Or acu = "C" Then

        Select Case "" & mytablex.Fields("tipodoc")

            Case "1", "A", "B", "C", "G", "D"

            Case "J", "K", "L", "M", "P"

            Case Else
                mytablex.Close
                Exit Function

        End Select

    End If

    racu = "" & mytablex.Fields("tipodoc")
    flagcruce = "" & mytablex.Fields("crucedoc")
    busca_tipo = 1

    If sw = 22 Then
        busca_tipo = 0

        If "" & mytablex.Fields("tipodoc") = "S" Or "" & mytablex.Fields("tipodoc") = "T" Then
            busca_tipo = 22

        End If

        Exit Function

    End If
   
    'Label16 = "" & mytablex.Fields("descripcio")
    If sw = 8 Then
        If "" & mytablex.Fields("obliga") = "S" Then
            busca_tipo = 8

        End If

    End If

    If sw = 7 Then
        If IsNumeric("" & Numero) Then
            'mytablex.Edit
            mytablex.Fields("numero") = "" & Numero
            mytablex.Update

        End If

    End If

    If sw = 9 Then
        sdx = Val("" & mytablex.Fields("numero")) + 1
        Numero = "" & sdx
        busca_tipo = 1

    End If

    If sw = 6 Then
        If Len(serie) = 0 Then
            'If tipoclie = "I" Then
            'serie = "I01"
            'Else
            serie = "" & mytablex.Fields("serie")

            ' End If
        End If

        busca_tipo = 1

    End If

    If sw = 2 Then
        flagcruce = "" & mytablex.Fields("crucedoc")

        If Len(bodega) = 0 Then
            bodega = "" & mytablex.Fields("bodega")

        End If

        busca_tipo = 1

    End If

    If sw = 1 Or sw = 0 Then
        flage = "" & mytablex.Fields("flage")
        racu = "" & mytablex.Fields("tipodoc")
        busca_tipo = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_tipo1(sw As Integer) As String

    Dim mytablex As New ADODB.Recordset

    'Label16 = ""
    mytablex.Open "SELECT * FROM tipo where  tipo='" & ttipo & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        If sw = 0 Then
            busca_tipo1 = "" & mytablex.Fields("te")

        End If

        If sw = 1 Then
            busca_tipo1 = "" & mytablex.Fields("ts")

        End If

        If sw = 2 Then
            bodega = "" & mytablex.Fields("bodega")

        End If

    End If

    mytablex.Close

End Function

Function busca_codigo()

    Dim mytablex As New ADODB.Recordset

    Label17 = ""

    If tipoclie = "P" Then
        mytablex.Open "SELECT * FROM proveedo where codigo='" & codigo & "'", cn, adOpenKeyset, adLockOptimistic

    End If

    If tipoclie = "C" Then
        mytablex.Open "SELECT * FROM clientes where codigo='" & codigo & "'", cn, adOpenKeyset, adLockOptimistic

    End If

    If tipoclie = "V" Then
        mytablex.Open "SELECT * FROM vendedor where codigo='" & codigo & "'", cn, adOpenKeyset, adLockOptimistic

    End If

    If mytablex.RecordCount = 0 Then 'si existe
        mytablex.Close
        Exit Function

    End If

    Label17 = "" & mytablex.Fields("nombre")

    If Len(moneda) = 0 Then
        moneda = "" & mytablex.Fields("moneda")

    End If

    If tipoclie <> "V" Then
      
        If Len(fpago) = 0 Then
            fpago = "" & mytablex.Fields("fpago")

        End If

        If Len(vendedor) = 0 Then
            vendedor = "" & mytablex.Fields("vendedor")

        End If

        If Len(dias) = 0 Then
            dias = "" & mytablex.Fields("diapago")

        End If

    End If

    If Len(moneda) = 0 Then
        moneda = "S"

    End If

    If Len(fpago) = 0 Then
        fpago = "1"

    End If

    If Val(dias) = 0 Then
        dias = "1"

    End If

    If Len(Trim("" & destino)) = 0 Then
        destino = "" & mytablex.Fields("direccion")

    End If
   
    busca_codigo = 1
    mytablex.Close

End Function

Function busca_vendedor()
    zona = ""

    Dim rsexiste As New ADODB.Recordset

    rsexiste.Open "SELECT * FROM vendedor where  codigo='" & vendedor & "'", cn, adOpenKeyset, adLockOptimistic

    If rsexiste.RecordCount > 0 Then  'si existe
        busca_vendedor = 1
        zona = "" & rsexiste.Fields("zona")

    End If

End Function

Function busca_local1(buf As String)

    Dim rsexiste As New ADODB.Recordset

    rsexiste.Open "SELECT * FROM tlocal where  codigo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If rsexiste.RecordCount > 0 Then  'si existe
        busca_local1 = 1

    End If

End Function

Function busca_transporte()

    Dim rsexiste As New ADODB.Recordset

    rsexiste.Open "SELECT * FROM transpor where  codigo='" & transporte & "'", cn, adOpenKeyset, adLockOptimistic

    If rsexiste.RecordCount > 0 Then  'si existe
        busca_transporte = 1

    End If

End Function

Function busca_fpago()

    Dim rsexiste As New ADODB.Recordset

    rsexiste.Open "SELECT * FROM fpago where  fpago='" & fpago & "'", cn, adOpenKeyset, adLockOptimistic

    If rsexiste.RecordCount > 0 Then  'si existe
        If moneda = "" & rsexiste.Fields("moneda") Then
            busca_fpago = 1

        End If

    Else
        MsgBox "Debe ser moneda=" & moneda, 48, "Aviso"

    End If

    rsexiste.Close

End Function

Function busca_bodega(buf0 As String, buf As String, sw As Integer)

    Dim mytablex As New ADODB.Recordset

    If sw = 0 Then

        'nbodega = ""
    End If

    If sw = 1 Then
        nbodega1 = ""

    End If

    mytablex.Open "SELECT * FROM bodega where  local='" & buf0 & "' and codigo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
      
        busca_bodega = 1

        If sw = 0 Then

            'nbodega = Mid$("" & mytablex.Fields("nombre"), 1, 10)
            'NBODEGA.Top = 150
            'NBODEGA.Height = 375
            'NBODEGA.Left = 9250
            'MsgBox "xxx"
        End If

        If sw = 1 Then
            nbodega1 = Mid$("" & mytablex.Fields("nombre"), 1, 10)

        End If

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Sub sql_detalle()

    Dim buf As String

    On Error GoTo cmd34_err

    buf = "select * from " & dgusuario & " ORDER BY hora "
    Data2.Connect = "foxpro 2.5;"
    Data2.DatabaseName = globaldat
    Data2.RecordSource = buf
    Data2.refresh
    dbgrid2.refresh
               
    'sql_importa
    'If Data2.Recordset.EOF = True And Data2.Recordset.BOF = True Then
    '   D ata2.Recordset.AddNew
    '   Data2.Recordset.Update
    'End If
    Exit Sub
cmd34_err:
    MsgBox "Error en select " & error$, 48, "Aviso"
    Exit Sub

End Sub

Function busca_producto(buf As String, sw As Integer, canti As Double)

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim xbuf     As String

    Dim found    As Integer

    Dim sw1      As Integer

    Dim ybuf     As String

    Dim buf1     As String

    Dim I        As Integer

    Dim ssw      As Integer

    Dim sfound   As String

    If Val(paridad) <= 0 Then
        paridad = "1"

    End If

    xbuf = buf
    sw1 = 0
    ybuf = ""
    sfound = "V"

    'If valida_flag("" & racu) = 2 Then    'compras
    If Mid$(buf, 1, 1) = "!" Then   'si es codigo de proveedor
        xbuf = Mid$(buf, 2, Len(buf) - 1)

        If Len(xbuf) = 0 Then
            Exit Function

        End If

        ybuf = xbuf
   
        found = busca_cod_prov("" & codigo, xbuf)

        If found = 0 Then
            Exit Function

        End If

        found = verifica_doble("" & xbuf)

        If found = 1 Then
            Exit Function

        End If

    End If

    sw = 0
    'validamos si es que tiene busqueda por codigo proveedor
    buf1 = xbuf

    '' 05/07/2018 Busqueda por Codigo de barras *
    If mytablex.State = 1 Then mytablex.Close
      
    If Mid(buf, 1, 1) = "*" Then
        buf = Mid(buf, 2, Len(buf) - 1)
        mytablex.Open "SELECT * FROM producto where barras='" & buf & "'", cn, adOpenStatic, adLockOptimistic
              
        If mytablex.RecordCount > 0 Then
            buf = "" & mytablex.Fields("producto")

        End If
              
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

    Else
      
        mytablex.Open "SELECT * FROM producto where producto='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.Close
            found = busca_equiva(buf1) 'busca en la table codigo barras

            If found = 0 Then
                Exit Function

            End If

            mytablex.Open "SELECT * FROM producto where producto='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic

            If mytablex.RecordCount = 0 Then
                mytablex.Close
                Exit Function

            End If

        End If

    End If
     
    '' 05/07/2018 Busqueda por Codigo de barras *
  
    'miramos si es compra
    'sfound = tipo_costo("" & ttipo)
    'If Trim(sfound) = "C" Or Trim(sfound) = "P" Or Len(Trim(sfound)) = 0 Then
    '   graba_temporald mytablex, mytabley, sw, sfound
    '   sw1 = 1
    '   busca_producto = 1
    '   mytablex.Close
    '   mytabley.Close
    '   Exit Function
    ' End If
         
    '-- ahora busca los precios
a12345:
    mytabley.Open "SELECT * FROM precios where  producto='" & "" & mytablex.Fields("producto") & "' and local='" & local1 & "'", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount = 0 Then  'si existe
        mytabley.AddNew
        mytabley.Fields("local") = Trim("" & local1)
        mytabley.Fields("producto") = Trim("" & mytablex.Fields("producto"))
        mytabley.Fields("unidad1") = "UND"
        mytabley.Fields("factor1") = 1
        mytabley.Update
        'MsgBox "No existe Precio venta en dicho Local ", 48, "Aviso"
        'mytablex.Close
        mytabley.Close
        GoTo a12345

    End If
         
    graba_temporald mytablex, mytabley, sw, canti
    sw1 = 1
    busca_producto = 1
    mytablex.Close
        
    'If sw1 = 1 And Len(ybuf) > 0 Then
    'If valida_flag("" & racu) = 2 Then    'compras
    '   found = crea_nuevos_proveedores("" & codigo, "" & xbuf, "" & ybuf)
    'End If
    'End If
    mytabley.Close

End Function

Sub graba_temporald(mytablex As ADODB.Recordset, _
                    mytabley As ADODB.Recordset, _
                    sw As Integer, _
                    canti As Double)

    Dim found    As Integer

    Dim pventa1  As Double

    Dim costou   As Double

    Dim buf      As String

    Dim mytables As New ADODB.Recordset

    ''''26/09/2017 kenyo Costos al por mayor Factura de Compra
    Dim factor   As Double

    factor = Val("" & mytablex.Fields("factor"))
    ''''26/09/2017 kenyo Costos al por mayor Factura de Compra

    pventa1 = Val("" & mytabley.Fields("pventa1"))

    ''''26/09/2017 kenyo Costos al por mayor Factura de Compra
    'costou = Val(Format(Val("" & mytablex.Fields("costou")), "0.00"))
    costou = Val(Format(Val("" & mytablex.Fields("costou")), "0.00000"))
    ''''26/09/2017 kenyo Costos al por mayor Factura de Compra

    If Val(paridad) <= 0 Then
        paridad = "1"

    End If

    'MsgBox moneda
    If "" & moneda = "S" Then
        If "" & mytablex.Fields("monedav") = "D" Then
            pventa1 = Val("" & mytabley.Fields("pventa1")) * Val(paridad)

        End If

        If "" & mytablex.Fields("monedaC") = "D" Then
            costou = costou * Val(paridad)

        End If

    End If

    If "" & moneda = "D" Then
        If "" & mytablex.Fields("monedav") = "S" Then
            pventa1 = Val("" & mytabley.Fields("pventa1")) / Val(paridad)

        End If

        If "" & mytablex.Fields("monedaC") = "S" Then
            costou = costou / Val(paridad)

        End If

    End If

    mytables.Open "SELECT * FROM DUENO where  local='" & local1 & "' and producto='" & "" & mytablex.Fields("producto") & "' ", cn, adOpenKeyset, adLockOptimistic

    If mytables.RecordCount > 0 Then  'si existe
        dbgrid2.columns("ccosto") = Trim("" & mytables.Fields("codigo"))

    End If

    mytables.Close

    ''''26/09/2017 kenyo Costos al por mayor Factura de Compra
    'dbgrid2.columns("proveedorp") = Format(costou, "0.00") 'costo anterior
    dbgrid2.columns("proveedorp") = Format(costou * factor, "0.00") 'costo anterior
    ''''26/09/2017 kenyo Costos al por mayor Factura de Compra

    dbgrid2.columns("producto") = "" & mytablex.Fields("producto")
    'dbGrid2.Columns("proveedorp") = "" '& mytablex.Fields("proveedor1")
    dbgrid2.columns("tipo") = "" & ttipo
    dbgrid2.columns("serie") = "" & serie
    dbgrid2.columns("numero") = "" & Numero
    dbgrid2.columns("vendedor") = "" & vendedor
    dbgrid2.columns("descripcio") = Mid$("" & mytablex.Fields("descripcio"), 1, 60)
    dbgrid2.columns("cantidad") = 1

    If canti > 0 Then
        dbgrid2.columns("cantidad") = canti

    End If

    dbgrid2.columns("unidad") = "" & mytabley.Fields("unidad1")
    dbgrid2.columns("factor") = Val("" & mytabley.Fields("factor1"))
    dbgrid2.columns("precio") = pventa1
    dbgrid2.columns("total") = pventa1
    dbgrid2.columns("subtotal") = pventa1
    dbgrid2.columns("descuento") = 0
    dbgrid2.columns("isc") = Val("" & mytablex.Fields("isc"))
    dbgrid2.columns("detraccion") = Val("" & mytablex.Fields("detraccion"))

    'DBGrid2.Columns(13) = Val("" & mytablex.Fields("tax"))
    If valida_flag("" & racu) = "2" Then  'compras
        dbgrid2.columns("unidad") = "" & mytablex.Fields("unidad")
        dbgrid2.columns("factor") = Val("" & mytablex.Fields("factor"))
        dbgrid2.columns("precio") = costou * Val("" & mytablex.Fields("factor"))
        dbgrid2.columns("total") = costou * Val("" & mytablex.Fields("factor"))
        dbgrid2.columns("subtotal") = costou * Val("" & mytablex.Fields("factor"))

    End If

    If valida_flag("" & racu) = "1" Then 'ventas
        dbgrid2.columns("unidad") = "" & mytabley.Fields("unidad1")
        dbgrid2.columns("factor") = Val("" & mytabley.Fields("factor1"))
        dbgrid2.columns("precio") = pventa1
        dbgrid2.columns("total") = pventa1
        dbgrid2.columns("subtotal") = pventa1

    End If

    buf = tipo_costo("" & ttipo)

    Select Case buf

        Case "V"
            dbgrid2.columns("precio") = pventa1

    End Select

    dbgrid2.columns("deslipo") = 0
    dbgrid2.columns("tax") = 0
    dbgrid2.columns("flete") = Val("" & mytablex.Fields("flete"))
    dbgrid2.columns("impuesto") = 0
    dbgrid2.columns("ivap") = Val("" & mytablex.Fields("ivap"))
    dbgrid2.columns("igv") = Val("" & mytablex.Fields("igv"))
    dbgrid2.columns("percepcion") = Val("" & mytablex.Fields("percepcion"))
    dbgrid2.columns("linea") = "" & mytablex.Fields("linea")

    dbgrid2.columns("descuento") = 0
    dbgrid2.columns("neto") = 0

    '---------pone a quien pertenece --------------------
    dbgrid2.columns("l1") = "" '& mytablex.Fields("c11")
    dbgrid2.columns("l2") = "" '& mytablex.Fields("c12")
    dbgrid2.columns("l3") = "" '& mytablex.Fields("c13")
    dbgrid2.columns("l4") = "" '& mytablex.Fields("c14")

    'LAS FAMILIAS+SUBFAMILIA+MARCA+SECCION
    dbgrid2.columns("familia") = "" & mytablex.Fields("Familia")
    dbgrid2.columns("subfamilia") = "" & mytablex.Fields("subFamilia")
    dbgrid2.columns("marca") = "" & mytablex.Fields("marca")
    'DBGrid2.columns("hora") = Format(hora, "hh:MM:ss")
    dbgrid2.columns("hora") = Format(Now, "hh:MM:ss")

    'If bodega = "01" Then
    '   found = ver_docena1(mytabley)
    'End If

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
    'If racu = "E" Or racu = "N" Then  'si es nota credito por compras o ventas
    '   If Val("" & DBGrid2.columns("precio")) >= 0 Then
    '      DBGrid2.columns("precio") = -Val("" & DBGrid2.columns("precio"))
    '   End If
    'End If
    '
    If racu = "E" Or racu = "N" Then  'si es nota credito por compras o ventas
        If Val("" & dbgrid2.columns("precio")) >= 0 Then
            dbgrid2.columns("precio") = Val("" & dbgrid2.columns("precio"))

        End If

    End If

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    '-----------------------------
    calcula_igv 0

End Sub

'' 11/12/2017 SubReceta

Function verifica_receta(buf As String)

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM receta where producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        verifica_receta = 1

    End If

    mytablex.Close

End Function

'' 11/12/2017 SubReceta

Sub suma_linea()

    Dim sdx As Double

    'sdx = Val("" & Data2.Recordset.Fields("precio")) * Val("" & Data2.Recordset.Fields("cantidad"))
    'Data2.Recordset.Fields("total") = Val(Format(sdx, "0.00"))
    'Data2.Recordset.Fields("neto") = Val(Format(sdx, "0.00"))
End Sub

Sub calcula_igv(sw As Integer)

    Dim sdx     As Double

    Dim sdx1    As Double

    Dim sdx2    As Double

    Dim tdscto  As Double

    Dim tdscto1 As Double

    Dim xtivap  As Double

    Dim found   As Integer

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    'If racu = "E" Or racu = "N" Then  'si es nota credito por compras o ventas
    '   If Val("" & DBGrid2.columns("precio")) >= 0 Then
    '      DBGrid2.columns("precio") = -Val("" & DBGrid2.columns("precio"))
    '      DBGrid2.columns("total") = Val("" & DBGrid2.columns("precio")) * Val("" & DBGrid2.columns("cantidad"))
    '   End If
    'End If
    If racu = "E" Or racu = "N" Then  'si es nota credito por compras o ventas
        If Val("" & dbgrid2.columns("precio")) >= 0 Then
            dbgrid2.columns("precio") = Val("" & dbgrid2.columns("precio"))
            dbgrid2.columns("total") = Val("" & dbgrid2.columns("precio")) * Val("" & dbgrid2.columns("cantidad"))

        End If

    End If

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    dbgrid2.columns("neto") = Val("" & dbgrid2.columns("cantidad")) * Val("" & dbgrid2.columns("precio"))
    dbgrid2.columns("descuento") = Val("" & dbgrid2.columns("neto")) * Val("" & dbgrid2.columns("deslipo")) / 100 + Val("" & dbgrid2.columns("neto")) * Val("" & dbgrid2.columns("destopo")) / 100
    dbgrid2.columns("total") = Val("" & dbgrid2.columns("neto")) - Val("" & dbgrid2.columns("descuento")) 'cobrar
    dbgrid2.columns("subtotal") = Val("" & dbgrid2.columns("total")) / (1 + Val("" & dbgrid2.columns("igv")) / 100) 'calcular descuento
    dbgrid2.columns("impuesto") = Val("" & dbgrid2.columns("total")) - Val("" & dbgrid2.columns("subtotal")) 'cobrar
    dbgrid2.columns("tivap") = Val("" & dbgrid2.columns("total")) * Val("" & dbgrid2.columns("ivap")) / 100
    dbgrid2.columns("tdetra") = Val("" & dbgrid2.columns("total")) * Val("" & dbgrid2.columns("detraccion")) / 100   'calcular descuento
    dbgrid2.columns("tpercepcio") = Val("" & dbgrid2.columns("total")) * Val("" & dbgrid2.columns("percepcion")) / 100   'calcular descuento

    If Trim(menup.Label10) = "ARGENTINA" Then
        dbgrid2.columns("tpercepcio") = Val("" & dbgrid2.columns("subtotal")) * Val("" & dbgrid2.columns("percepcion")) / 100   'calcular descuento

    End If

    dbgrid2.columns("total") = Val("" & dbgrid2.columns("total")) + Val("" & dbgrid2.columns("tpercepcio")) - Val("" & dbgrid2.columns("tdetra")) 'cobrar
    dbgrid2.columns("servicioco") = Val("" & dbgrid2.columns("subtotal")) * Val("" & dbgrid2.columns("serviciopo")) / 100      'calcular descuento
    dbgrid2.columns("tisc") = Val("" & dbgrid2.columns("subtotal")) * Val("" & dbgrid2.columns("isc")) / 100

End Sub

Sub consulta_producto(buf As String)
    cerrar_data1
    sw_consulta = 0
    Combo1.Clear
    Check1.Value = 0
    Check1.Visible = False
    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "Producto.Descripcio"
    Combo2.AddItem "Producto.Producto"
    Combo2.AddItem "Producto.Marca"
    Combo2.AddItem "Producto.Familia"
    Combo2.AddItem "Producto.Subfamilia"
    Combo2.AddItem "Producto.barras"
    Combo2.AddItem "precios.Unidad1"
    Combo2.ListIndex = 0

    Combo1.AddItem "Producto.Descripcio"
    Combo1.AddItem "Producto.Producto"
    Combo1.AddItem "Producto.Marca"
    Combo1.AddItem "Producto.Familia"
    Combo1.AddItem "Producto.Subfamilia"
    Combo1.AddItem "Producto.barras"
    Combo1.AddItem "precios.Unidad1"
    'Combo1.AddItem "Producto.proveedor1"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = "" & buf
    buffer.SetFocus
    opcion1 = "8"
    'If valida_flag("" & racu) = 1 Then    'compras
    Check1.Visible = True
    'Check1.Value = 1
    '   opcion1 = "45"
    'End If

    dbgrid2.Enabled = False

    If Len(Trim(buf)) > 0 Then
        Command1_Click
        Exit Sub

    End If

    Set DBGrid1.DataSource = Nothing
    buffer.SetFocus

End Sub

Sub consulta_productose(buf As String)

    sw_consulta = 0
    Combo1.Clear
    Check1.Value = 0
    Check1.Visible = False
    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "Producto.Descripcio"
    Combo2.AddItem "Producto.Producto"
    Combo2.AddItem "Producto.Marca"
    Combo2.AddItem "Producto.Familia"
    Combo2.AddItem "Producto.Subfamilia"
    Combo2.AddItem "Producto.barras"
    Combo2.AddItem "precios.Unidad1"
    Combo2.ListIndex = 0

    Combo1.AddItem "Producto.Descripcio"
    Combo1.AddItem "Producto.Producto"
    Combo1.AddItem "Producto.Marca"
    Combo1.AddItem "Producto.Familia"
    Combo1.AddItem "Producto.Subfamilia"
    Combo1.AddItem "Producto.barras"
    Combo1.AddItem "precios.Unidad1"
    'Combo1.AddItem "Producto.proveedor1"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = "" & buf
    buffer.SetFocus
    opcion1 = "888"
    'If valida_flag("" & racu) = 1 Then    'compras
    Check1.Visible = True
    'Check1.Value = 1
    '   opcion1 = "45"
    'End If
    dbgrid2.Enabled = False

    If Len(Trim(buf)) > 0 Then
        Command1_Click
        Exit Sub

    End If

    Set DBGrid1.DataSource = Nothing
    buffer.SetFocus

End Sub

Sub consulta_rproducto()
    cerrar_data1
    sw_consulta = 0
    Combo1.Clear

    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "Descripcio"
    Combo2.AddItem "Producto"
    Combo2.AddItem "Familia"
    Combo2.AddItem "Seccion"
    Combo2.AddItem "Categoria"
    Combo2.AddItem "Marca"
    'Combo2.AddItem "proveedor1"
    Combo2.ListIndex = 0
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "Producto"
    Combo1.AddItem "Familia"
    Combo1.AddItem "Seccion"
    Combo1.AddItem "Categoria"
    Combo1.AddItem "Marca"
    'Combo1.AddItem "proveedor1"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "50"

    If Len(Trim(buffer)) > 0 Then
        Command1_Click
        Exit Sub

    End If

    Set dbgrid3.DataSource = Nothing
    buffer.SetFocus

End Sub

Sub cerrar_data2()

    On Error GoTo cmd4_err

    Data2.Recordset.Close
    Exit Sub
cmd4_err:
    Exit Sub

End Sub

Function suma_grid2()

    Dim fila As Integer

    Dim suma As Double

    suma = 0

    For fila = 0 To Data2.Recordset.RecordCount - 1
        dbgrid2.Row = fila    'El índice de la primera fila empieza en 0.
        suma = suma + Val("" & dbgrid2.columns("descripcio").Value)
    Next

End Function

Sub borrar_detalle_todo_registro()

    On Error GoTo cmd45_err

    ir_primero
amk12:
    Data2.Recordset.Delete
    Data2.refresh
    GoTo amk12
    Exit Sub
cmd45_err:
    Exit Sub

End Sub

Sub borrar_detalle_linea()
    Data2.Recordset.Delete
    dbgrid2.refresh

End Sub

Sub ir_ultimo()

    On Error GoTo cmd50_err

    sumar_detalle
    Data2.Recordset.MoveLast
    Exit Sub
cmd50_err:
    Exit Sub

End Sub

Sub ir_primero()

    On Error GoTo cmd51_err

    Data2.Recordset.MoveFirst
    Exit Sub
cmd51_err:
    Exit Sub

End Sub

Sub solo_ir_ultimo()

    On Error GoTo cmd53_err

    Data2.Recordset.MoveFirst
    Exit Sub
cmd53_err:
    Exit Sub

End Sub

Sub cerrar_data1()
    'On Error GoTo cmd17_err
    'Data1.Recordset.Close
    'Exit Sub
    'cmd17_err:
    Exit Sub

End Sub

Sub sumar_detalle2()

    On Error GoTo cmd34_err

    Dim fila       As Integer

    Dim xtotal     As Double

    Dim xdescuento As Double

    Dim xneto      As Double

    Dim ximpuesto  As Double

    Dim xsubtotal  As Double

    Dim xc1        As Double

    Dim xc2        As Double

    Dim xc3        As Double

    Dim xc4        As Double

    Dim xgravado   As Double

    Dim vr

    Dim xntcant As Double

    xntcant = 0
    xc1 = 0
    xc2 = 0
    xc3 = 0
    xc4 = 0
    xgravado = 0
    xtotal = 0
    xdescuento = 0
    xneto = 0
    ximpuesto = 0
    xsubtotal = 0

    'dbrecords = Data2.Recordset.RecordCount
    'For fila = 0 To DBGrid2.ApproxCount - 1
    For fila = 0 To Data2.Recordset.RecordCount - 1
        dbgrid2.Row = fila
        'If "" & dbgrid2.columns(34).Value = "1" Then
        '   xc1 = xc1 + Val("" & dbgrid2.columns("total").Value)
        'End If
        'If "" & dbgrid2.columns(35).Value = "1" Then
        '   xc2 = xc2 + Val("" & dbgrid2.columns("total").Value)
        'End If
        'If "" & dbgrid2.columns(36).Value = "1" Then
        '   xc3 = xc3 + Val("" & dbgrid2.columns("total").Value)
        'End If
        'If "" & dbgrid2.columns(37).Value = "1" Then
        '   xc4 = xc4 + Val("" & dbgrid2.columns("total").Value)
        'End If
        xntcant = xntcant + Val("" & dbgrid2.columns("cantidad").Value) 'suma bruto
        xneto = xneto + Val("" & dbgrid2.columns("neto").Value) 'suma bruto
        xdescuento = xdescuento + Val("" & dbgrid2.columns("descuento").Value) 'suma descuento
        xsubtotal = xsubtotal + Val("" & dbgrid2.columns("subtotal").Value) ' suma subtotal
        ximpuesto = ximpuesto + Val("" & dbgrid2.columns("impuesto").Value) 'suma impuesto
        xtotal = xtotal + Val("" & dbgrid2.columns("total").Value)  'suma total
    Next
    ntcant = Format(xntcant, "0.00")
    txneto = Format(xneto, "0.00")
    txdescuento = Format(xdescuento, "0.00")
    txsubtotal = Format(xsubtotal, "0.00")
    tximpuesto = Format(ximpuesto, "0.00")
    txtotal = Format(xtotal, "0.00")
    c1 = Format(xc1, "0.00")
    c2 = Format(xc2, "0.00")
    c3 = Format(xc3, "0.00")
    c4 = Format(xc4, "0.00")
    Exit Sub
cmd34_err:
    MsgBox "Error " & error$ & " " & fila, 24, "Aviso"
    Exit Sub

End Sub

Sub sumar_detalle()

    On Error GoTo cmd35_err

    Dim fila        As Integer

    Dim xtotal      As Double

    Dim xdescuento  As Double

    Dim xneto       As Double

    Dim ximpuesto   As Double

    Dim xsubtotal   As Double

    Dim xflete      As Double

    Dim sdx         As Double

    Dim xc1         As Double

    Dim xc2         As Double

    Dim xc3         As Double

    Dim xc4         As Double

    Dim xc5         As Double

    Dim xc6         As Double

    Dim xc7         As Double

    Dim xc8         As Double

    Dim xc9         As Double

    Dim xpercep     As Double

    Dim xcostopais  As Double

    Dim xgravado    As Double

    Dim xisc        As Double

    Dim xdetraccion As Double

    Dim xivap       As Double

    Dim vr

    Dim xntcant As Double

    xisc = 0
    xivap = 0
    xdetraccion = 0
    xpercep = 0
    xgravado = 0
    xntcant = 0
    xcostopais = 0
    xc1 = 0
    xc2 = 0
    xc3 = 0
    xc4 = 0
    xc5 = 0
    xc6 = 0
    xc7 = 0
    xc8 = 0
    xc9 = 0

    xtotal = 0
    xdescuento = 0
    xneto = 0
    ximpuesto = 0
    xsubtotal = 0
    xflete = 0
    'dbrecords = Data2.Recordset.RecordCount
    'For fila = 0 To DBGrid2.ApproxCount - 1

    'sumar el total
    'sumar la cantidad
    'dividir total/cantidad lo que toca a cada uno
    'poner en el campo flete
    'este dato sumar al costo
    Data2.Recordset.MoveFirst
    Do

        If Data2.Recordset.EOF Then Exit Do
        Data2.Recordset.Edit
        resuma_precios 0
        Data2.Recordset.Update

        If Val("" & Data2.Recordset.Fields("igv")) = 0 Then
            xgravado = xgravado + Val("" & Data2.Recordset.Fields("total"))

        End If

        xivap = xivap + Val("" & Data2.Recordset.Fields("tivap"))
        xflete = xflete + Val("" & Data2.Recordset.Fields("flete"))  'flete
        xntcant = xntcant + Val("" & Data2.Recordset.Fields("cantidad")) 'suma bruto
        xtotal = xtotal + Val("" & Data2.Recordset.Fields("total"))
        xdescuento = xdescuento + Val("" & Data2.Recordset.Fields("descuento"))
        xneto = xneto + Val("" & Data2.Recordset.Fields("neto"))
        ximpuesto = ximpuesto + Val("" & Data2.Recordset.Fields("impuesto"))
        xsubtotal = xsubtotal + Val("" & Data2.Recordset.Fields("subtotal"))
        xpercep = xpercep + Val("" & Data2.Recordset.Fields("tpercepcio"))
        xdetraccion = xdetraccion + Val("" & Data2.Recordset.Fields("tdetra"))
        xisc = xisc + Val("" & Data2.Recordset.Fields("tisc"))
        xcostopais = xcostopais + Val("" & Data2.Recordset.Fields("cantdev")) * Val("" & Data2.Recordset.Fields("cantidad"))
        Data2.Recordset.MoveNext
    Loop
    'calcular el flete
    costopais = Format(xcostopais, "0.00")
    tflete = Format(xflete, "0.00")
    txisc = Format(xisc, "0.00")
    txivap = Format(xivap, "0.00")
    txdetraccion = Format(xdetraccion, "0.00")
    gravado = Format(xgravado, "0.00")
    ntcant = Format(xntcant, "0.00")

    'txtotal = Format(xtotal, "0.00")
    txtotal = Format(xtotal, "0.00")
    txpercepcio = Format(xpercep, "0.00")
    txdescuento = Format(xdescuento, "0.00")
    txneto = Format(xneto, "0.00")
    tximpuesto = Format(ximpuesto, "0.00")
    txsubtotal = Format(xsubtotal, "0.00")

    c1 = Format(xc1, "0.00")
    c2 = Format(xc2, "0.00")
    c3 = Format(xc3, "0.00")
    c4 = Format(xc4, "0.00")
    c5 = Format(xc5, "0.00")
    c6 = Format(xc6, "0.00")
    c7 = Format(xc7, "0.00")
    c8 = Format(xc8, "0.00")
    c9 = Format(xc9, "0.00")
    Exit Sub
cmd35_err:
    'MsgBox "Error " & Error$ & " " & fila, 24, "Aviso"
    Exit Sub

End Sub

Sub habilita_cabeza(sw As Integer)

    Dim xsw As Variant

    If sw = 0 Then
        xsw = True

    End If

    If sw = 1 Then
        xsw = False

    End If

    image1.Enabled = xsw
    codigo.Enabled = xsw
    partida.Enabled = xsw
    destino.Enabled = xsw
    fecha.Enabled = xsw
    fechae.Enabled = xsw
    moneda.Enabled = xsw
    vendedor.Enabled = xsw
    fpago.Enabled = xsw
    transporte.Enabled = xsw
    paridad.Enabled = xsw
    dias.Enabled = xsw
    bodega.Enabled = xsw
    'localf.Enabled = xsw
    'bodegaf.Enabled = xsw
    observa.Enabled = xsw
    'estado.Enabled = xsw

End Sub

Sub habilita_detalle(sw As Integer)

    Dim xsw As Variant

    If sw = 0 Then
        xsw = True

    End If

    If sw = 1 Then
        xsw = False

    End If

    dbgrid2.Enabled = xsw

End Sub

Sub habilita_numero(sw As Integer)

    Dim xsw As Variant

    If sw = 0 Then
        xsw = True

    End If

    If sw = 1 Then
        xsw = False

    End If

    'local1.Enabled = xsw
    ttipo.Enabled = xsw
    serie.Enabled = xsw
    Numero.Enabled = xsw

End Sub

Function cargar_registrod()

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    ''' 11/12/2017 SubReceta
    'S es para guia de salida
    'mytablex.Open "SELECT * FROM " & dgusuariog & " where local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & numero & "'", cn, adOpenKeyset, adLockOptimistic
    If racu = "S" Or racu = "T" Or racu = "A" Or racu = "B" Or racu = "C" Or racu = "D" Or racu = "G" Then
        mytablex.Open "SELECT * FROM " & dgusuariog & " where  DUA IS NULL AND local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & Numero & "'", cn, adOpenKeyset, adLockOptimistic
    Else
        mytablex.Open "SELECT * FROM " & dgusuariog & " where  local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & Numero & "'", cn, adOpenKeyset, adLockOptimistic

    End If

    ''' 11/12/2017 SubReceta
    
    If mytablex.RecordCount = 0 Then 'si existe
        mytablex.Close
        Exit Function

    End If

    Do

        If mytablex.EOF Then Exit Do
        Data2.Recordset.AddNew

        For I = 0 To mytablex.Fields.count - 2
            Data2.Recordset.Fields(I) = mytablex.Fields(I)
        Next I

        Data2.Recordset.Fields("local") = "" & local1
        Data2.Recordset.Fields("tipo") = "" & ttipo
        Data2.Recordset.Fields("serie") = "" & serie
        Data2.Recordset.Fields("numero") = "" & Numero
        Data2.Recordset.Fields("vendedor") = "" & vendedor
        Data2.Recordset.Fields("moneda") = "" & moneda
        Data2.Recordset.Fields("bodega") = "" & bodega
        Data2.Recordset.Fields("localf") = "" & localf
        Data2.Recordset.Fields("bodegaf") = "" & bodegaf
        Data2.Recordset.Fields("acu") = "" & racu
        Data2.Recordset.Fields("flage") = "" & flage
        Data2.Recordset.Fields("tipoclie") = tipoclie
        Data2.Recordset.Update
        mytablex.MoveNext
    Loop
    '------------------------------------- ------------
    mytablex.Close

End Function

Function cargar_registrose()

    Dim I        As Integer

    Dim mytabley As Table

    Dim mytablex As New ADODB.Recordset

    If acu <> "3" Then Exit Function  'si no es servicio tecnico salir
    mytablex.Open "SELECT * FROM serviciotecnico where  local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & Numero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then 'si existe
        mytablex.Close
        Exit Function

    End If

    Set mytabley = mydbxglo.OpenTable(sgusuario)
    mytabley.Index = "servicio"
hyu1:
    mytabley.Seek "=", local1, ttipo, serie, Numero

    If Not mytabley.NoMatch Then
        mytabley.Delete
        GoTo hyu1

    End If

    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 2
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Fields("local") = "" & local1
        mytabley.Fields("tipo") = "" & ttipo
        mytabley.Fields("serie") = "" & serie
        mytabley.Fields("numero") = "" & Numero
        mytabley.Update
        mytablex.MoveNext
    Loop
    mytabley.Close
    '------------------------------------- ------------
    mytablex.Close

End Function

Sub proceso_impresion1()

    Dim found    As Integer

    Dim archivot As String

    On Error GoTo cmd6_err:

    cerrar_archivo
    factura_formato "" & local1, "" & ttipo, "" & serie, "" & Numero, "", 0
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    Exit Sub
cmd6_err:
    MsgBox "Mensaje, Error al iniciar Impresion " & error$
    Exit Sub

End Sub

Function verifica_doble(buf As String)

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    mytablex.Open "SELECT * FROM tipo where tipo='" & ttipo & "' and repitencia='S'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then  'si existe
        mytablex.Close
        Exit Function

    End If

    mytablex.Close

    Set mytabley = mydbxglo.OpenTable(dgusuario)
    mytabley.Index = "cuerpo"
    mytabley.Seek "=", ttipo, serie, Numero, buf

    If Not mytabley.NoMatch Then
        verifica_doble = 1 'estab esto

        'verifica_doble = 0
    End If

    mytabley.Close

End Function

Sub grabar_cuentaxc()

    Dim mytabley As New ADODB.Recordset

    Dim buf      As String

    On Error GoTo cmd2340_err

    '---------- validando si es cuenta corriente
    If valida_flag("" & racu) = 2 Then    'compras
        buf = "cuentap"
   
    End If

    If valida_flag("" & racu) = 1 Then
        buf = "cuentac"
   
    End If

    mytabley.Open "SELECT * FROM " & buf & " where  local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & Numero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount = 0 Then  'si existe
        mytabley.AddNew
        grabar_registro_cuentac mytabley
        mytabley.Update
    Else
        'mytabley.Edit
        grabar_registro_cuentac mytabley
        mytabley.Update

    End If

    mytabley.Close
    Exit Sub
cmd2340_err:
    MsgBox "Aviso en grabar cuentaxc " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub grabar_registro_cuentac(mytabley As ADODB.Recordset)

    Dim wfecha As String

    mytabley.Fields("fpago") = "" & busca_fpagoc("" & fpago)
    mytabley.Fields("zona") = "" & zona
    mytabley.Fields("grupo") = "C"
    mytabley.Fields("acu") = "" & acu
    mytabley.Fields("local") = "" & local1
    mytabley.Fields("tipo") = "" & ttipo
    mytabley.Fields("serie") = "" & serie
    mytabley.Fields("nombre") = Mid$("" & Label17, 1, 35)
    mytabley.Fields("vendedor") = "" & vendedor
    mytabley.Fields("numero") = "" & Numero
    mytabley.Fields("tipoclie") = "" & tipoclie
    mytabley.Fields("codigo") = "" & codigo
    mytabley.Fields("cuota") = "1"
    mytabley.Fields("fecha") = Format("" & fecha, "dd/mm/yyyy")
    mytabley.Fields("moneda") = "" & moneda
    mytabley.Fields("total") = Val("" & txtotal)
    mytabley.Fields("abono") = 0
    mytabley.Fields("saldo") = Val("" & txtotal)
    mytabley.Fields("estado") = "0"
   
    '' 27/11/2017 Correcion duplicidad de guia salida venta en creditos
    mytabley.Fields("dias") = Val("" & dias)
    '' 27/11/2017 Correcion duplicidad de guia salida venta en creditos
   
    wfecha = Format((CVDate(fecha) + Int("" & dias)), "dd/mm/yyyy")
    mytabley.Fields("fechav") = Format(wfecha, "dd/mm/yyyy")
    mytabley.Fields("c1") = Val("" & c1)
    mytabley.Fields("c2") = Val("" & c2)
    mytabley.Fields("c3") = Val("" & c3)
    mytabley.Fields("c4") = Val("" & c4)
    mytabley.Fields("c5") = Val("" & c5)
    mytabley.Fields("c6") = Val("" & c6)
    mytabley.Fields("c7") = Val("" & c7)
    mytabley.Fields("c8") = Val("" & c8)
    mytabley.Fields("c9") = Val("" & c9)
    mytabley.Fields("usuario") = "" & gusuario

    If Len(cajero) > 0 Then
        mytabley.Fields("usuario") = "" & cajero

    End If

    If Len(caja) = 0 Then
        caja = "00"

    End If

    mytabley.Fields("caja") = "" & caja
   
End Sub

Function busca_fpagoc(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM fpago where  fpago='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        busca_fpagoc = "" & mytablex.Fields("tipo")

    End If

    mytablex.Close

End Function

Function graba_fpagov()

    Dim mytabley As New ADODB.Recordset

    Dim mytablex As New ADODB.Recordset

    Dim xyfpago  As String

    '---------- validando si es cuenta corriente
    xyfpago = ""

    mytablex.Open "SELECT * FROM fpago where  fpago='" & fpago & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        xyfpago = "" & mytablex.Fields("tipo")

    End If

    mytabley.Open "SELECT * FROM fpagov where  local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & Numero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount = 0 Then 'si existe
        mytabley.AddNew
        grabar_registro_fpagov mytabley
        mytabley.Fields("acufp") = xyfpago
        mytabley.Update
    Else
        'mytabley.Edit
        grabar_registro_fpagov mytabley
        mytabley.Fields("acufp") = xyfpago
        mytabley.Update

    End If

    mytabley.Close
    mytablex.Close

End Function

Sub grabar_registro_fpagov(mytabley As ADODB.Recordset)
    mytabley.Fields("local") = "" & local1
    mytabley.Fields("tipo") = "" & ttipo
    mytabley.Fields("serie") = "" & serie
    mytabley.Fields("numero") = "" & Numero
    mytabley.Fields("tipoclie") = "" & tipoclie
    mytabley.Fields("codigo") = "" & codigo
    mytabley.Fields("fecha") = Format("" & fecha, "dd/mm/yyyy")
    mytabley.Fields("moneda") = "" & moneda
   
    mytabley.Fields("total") = Val("" & txtotal)
    mytabley.Fields("recibe") = Val("" & txtotal)
   
    ' Testing Proyecto Facturacion Electronica Nota de Credito 06/2018
    'If acu = "E" Then
    '   mytabley.Fields("total") = Val("" & txtotal)
    '   mytabley.Fields("recibe") = Val("" & txtotal)
    'End If
    ' Testing Proyecto Facturacion Electronica Nota de Credito 06/2018
   
    mytabley.Fields("usuario") = "" & gusuario

    If Len(cajero) > 0 Then
        mytabley.Fields("usuario") = "" & cajero

    End If

    mytabley.Fields("fpago") = "" & fpago
   
    '''09/10/2017 kenyo Testing Reportes

    mytabley.Fields("descripcio") = "" & busca_fpagoComprobante(fpago)
    '''09/10/2017 kenyo Testing Reportes
   
    mytabley.Fields("acu") = "" & racu
    mytabley.Fields("local") = local1 'globalocal
    mytabley.Fields("estado") = "2"
    mytabley.Fields("caja") = caja

    If Len(caja) = 0 Then
        mytabley.Fields("caja") = "00"

    End If

    mytabley.Fields("servicio") = Servicio
    mytabley.Fields("turno") = turno
    mytabley.Fields("vendedor") = vendedor

End Sub

'''09/10/2017 kenyo Testing Reportes
Function busca_fpagoComprobante(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT descripcio FROM fpago where  fpago='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        busca_fpagoComprobante = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close

End Function

'''09/10/2017 kenyo Testing Reportes

Sub generar_traslados()

End Sub

Function busca_linea(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM linea where  linea='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
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

    mytablex.Close

End Function

Sub ingreso_tallas(buf As String)

    Dim found As Integer

    linea = buf
    found = busca_linea(buf)

    If found = 0 Then Exit Sub
    pone_tallas
    Frame2.Visible = True
    t1.SetFocus

End Sub

Sub menu_carga()

    Dim found As Integer

    If Len(tipo1) = 0 Then Exit Sub
    If Len(serie1) = 0 Then Exit Sub
    If Len(numero1) = 0 Then Exit Sub

    found = busca_tipo_carga("" & tipo1)

    If found = 0 Then Exit Sub
    cargar_cotizaciones local1, tipo1, serie1, numero1
    cargar_cotizaciones local1, tipo1, serie2, numero2
    cargar_cotizaciones local1, tipo1, serie3, numero3
    cargar_cotizaciones local1, tipo1, serie4, numero4
    cargar_cotizaciones local1, tipo1, serie5, numero5
    cargar_cotizaciones local1, tipo1, serie6, numero6
    cargar_cotizaciones local1, tipo1, serie7, numero7
    sumar_detalle

End Sub

Sub cargar_cotizaciones(xlocal1 As String, _
                        xtipo1 As String, _
                        xserie1 As String, _
                        xnumero1 As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM " & xarchivo1 & " where  local='" & xlocal1 & "' and tipo='" & xtipo1 & "' and serie='" & xserie1 & "' and numero='" & xnumero1 & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then  'si existe
        mytablex.Close
        Exit Sub

    End If

    Do

        If mytablex.EOF Then Exit Do
        graba_archivo_detalle mytablex
        mytablex.MoveNext
    Loop
    mytablex.Close

End Sub

Sub graba_archivo_detalle(mytablex As ADODB.Recordset)

    Dim I As Integer

    Data2.Recordset.AddNew

    For I = 0 To mytablex.Fields.count - 1
        Data2.Recordset.Fields(I) = mytablex.Fields(I)
    Next I

    Data2.Recordset.Fields("tipo") = "" & ttipo
    Data2.Recordset.Fields("serie") = "" & serie
    Data2.Recordset.Fields("numero") = "" & Numero
    Data2.Recordset.Fields("vendedor") = "" & vendedor
    Data2.Recordset.Fields("moneda") = "" & moneda
    Data2.Recordset.Fields("bodega") = "" & bodega
    Data2.Recordset.Fields("localf") = "" & localf
    Data2.Recordset.Fields("bodegaf") = "" & bodegaf
    Data2.Recordset.Fields("acu") = "" & racu
    Data2.Recordset.Fields("flage") = "" & flage
    Data2.Recordset.Fields("local") = local1 '"" & globalocal
    Data2.Recordset.Fields("tipoclie") = tipoclie
         
    Data2.Recordset.Update

End Sub

Function busca_tipo_carga(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM tipo where   tipo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        busca_tipo_carga = 1

        Select Case "" & mytablex.Fields("tipodoc")

            Case "1", "A", "B", "C", "D", "G", "E", "F" 'VENTAS
                xarchivo = "FACTURA"
                xarchivo1 = "DETALLE"

            Case "J", "K", "L", "M", "P", "N", "O"  'COMPRAS
                xarchivo = "FACTURA"
                xarchivo1 = "DETALLE"

            Case "H"  'COTIZACION VENTAS
                xarchivo = "CCOTIZAV"
                xarchivo1 = "DCOTIZAV"

            Case "I"  'PEDIDO VENTAS
                xarchivo = "CPEDIDOV"
                xarchivo1 = "DPEDIDOV"

            Case "Q"  'REQUISICION COMPRAS
                xarchivo = "CREQUISA"
                xarchivo1 = "DREQUISA"

            Case "R"  'ORDEN COMPRA
                xarchivo = "CORDENC"
                xarchivo1 = "DORDENC"

            Case "T", "S" 'GUIA REMISION
                xarchivo = "FACTURA"
                xarchivo1 = "DETALLE"

        End Select

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Sub consulta_detalles()

    Dim rconsulta As New ADODB.Recordset

    Dim found     As Integer

    Dim buf       As String

    found = busca_tipo_carga("" & DBGrid1.columns(0))

    If found = 0 Then Exit Sub
    buf = "select Producto,Descripcio,Unidad,Factor,Cantidad,Precio,Total,Moneda from " & xarchivo1 & " where tipo='" & DBGrid1.columns(0) & "' and serie='" & DBGrid1.columns(1) & "' and numero='" & DBGrid1.columns(2) & "'"

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open buf, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True And rconsulta.BOF = True Then
        rconsulta.Close
        Exit Sub

    End If

    Set dbgrid3.DataSource = rconsulta
    dbgrid3.Visible = True
    dbgrid3.SetFocus

End Sub

Sub cerrar_dbgrid3()
    dbgrid3.Visible = False
    DBGrid1.SetFocus

End Sub

Sub pone_tallas()
    t1 = "" & dbgrid2.columns("t1")
    t2 = "" & dbgrid2.columns("t2")
    t3 = "" & dbgrid2.columns("t3")
    t4 = "" & dbgrid2.columns("t4")
    t5 = "" & dbgrid2.columns("t5")
    t6 = "" & dbgrid2.columns("t6")
    t7 = "" & dbgrid2.columns("t7")
    t8 = "" & dbgrid2.columns("t8")
    t9 = "" & dbgrid2.columns("t9")
    t10 = "" & dbgrid2.columns("t10")
    t11 = "" & dbgrid2.columns("t11")
    t12 = "" & dbgrid2.columns("t12")
    t13 = "" & dbgrid2.columns("t13")
    t14 = "" & dbgrid2.columns("t14")
    t15 = "" & dbgrid2.columns("t15")
    t16 = "" & dbgrid2.columns("t16")

End Sub

Sub decarga_saldo_talla(mytablex As ADODB.Recordset, _
                        mytabley As ADODB.Recordset, _
                        signo As Double)

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

Sub xxpone_locales()

    Dim found As Integer

    Observa1 = "" & dbgrid2.columns("observa1")
    observa2 = "" & dbgrid2.columns("observa2")
    observa3 = "" & dbgrid2.columns("observa3")
    observa4 = "" & dbgrid2.columns("observa4")

End Sub

Sub ingreso_locales()
    xxpone_locales
    Frame3.Visible = True
    Observa1.SetFocus

    'If acu = "R" Then 'si no es orden de compra
    '   l1.Enabled = False
    '   l2.Enabled = False
    '   l3.Enabled = False
    '   l4.Enabled = False
    'End If
    'l1.SetFocus
End Sub

Sub consulta_documento()
    cerrar_data1
    sw_consulta = 0
    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "Tipo"
    Combo2.AddItem "Numero"
    Combo2.AddItem "Codigo"
    Combo2.ListIndex = 0

    Combo1.Clear
    Combo1.AddItem "Tipo"
    Combo1.AddItem "Numero"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "22"
    Command1_Click

End Sub

Sub calcula_igv1()

    Dim sdx  As Double

    Dim sdx1 As Double

    Dim sdx2 As Double

    If racu = "E" Or racu = "N" Then   'si es nota credito compras o ventas
        If Val("" & Data2.Recordset.Fields("precio")) > 0 Then
            Data2.Recordset.Fields("precio") = -Val("" & Data2.Recordset.Fields("precio"))

        End If

    End If

    sdx = Val("" & Data2.Recordset.Fields("precio")) * Val("" & Data2.Recordset.Fields("cantidad"))
    Data2.Recordset.Fields("total") = Val(Format(sdx, "0.00"))  'total
    Data2.Recordset.Fields("neto") = Val(Format(sdx, "0.00"))  'neto
    sdx = Val("" & Data2.Recordset.Fields("neto")) * Val("" & Data2.Recordset.Fields("deslipo")) / 100
    sdx2 = Val("" & Data2.Recordset.Fields("neto")) - sdx
    Data2.Recordset.Fields("descuento") = Val(Format(sdx, "0.00"))  'descuento
    Data2.Recordset.Fields("total") = Val(Format(sdx2, "0.00"))  'total
    Data2.Recordset.Fields("subtotal") = 0
    Data2.Recordset.Fields("impuesto") = 0

    If Val("" & Data2.Recordset.Fields("total")) > 0 And Val("" & Data2.Recordset.Fields("igv")) > 0 Then
        sdx1 = 1 + Val("" & Data2.Recordset.Fields("igv")) / 100
        sdx1 = Val(Format(sdx1, "0.00"))
        sdx1 = Val("" & Data2.Recordset.Fields("total")) / sdx1
        Data2.Recordset.Fields("subtotal") = Val(Format(sdx1, "0.00"))  'subtotal
        sdx = Val("" & Data2.Recordset.Fields("total")) - Val("" & Data2.Recordset.Fields("subtotal"))
        Data2.Recordset.Fields("impuesto") = Val(Format(sdx, "0.00"))  'total

    End If

End Sub

Sub carga_dbgrid4()

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim sw       As Integer

    Dim xbodega  As String

    Dim xsaldo   As Double

    Dim xbuf     As String

    Dim xcosto   As Double

    Dim xcostou  As Double

    Dim xfactor  As Double

    Dim xunidad  As String

    Dim xmargen  As Double

    On Error GoTo cmd89012_err

    For I = 0 To 9
        campo_precios(I).unidad = ""
        campo_precios(I).factor = ""
        campo_precios(I).precio = ""
        campo_precios(I).costo = ""
        campo_precios(I).margen = ""
        campo_precios(I).stock = ""
    Next I

    xcostou = 0
    xunidad = "UND"
    xfactor = 1
    xbodega = bodega
    xsaldo = 0
    xcosto = 0
    sw = 0

    mytabley.Open "SELECT * FROM almacen where  local='" & local1 & "' and producto='" & xproducto & "' and bodega='" & xbodega & "'", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount > 0 Then  'si existe
        xsaldo = Val("" & mytabley.Fields("saldo"))

    End If

    mytabley.Close
    mytablex.Open "SELECT * FROM producto where  producto='" & xproducto & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        xcostou = Val(Format(Val("" & mytablex.Fields("costou")) * Val("" & mytablex.Fields("factor")), "0.00"))
        xfactor = Val("" & mytablex.Fields("factor"))
        xunidad = "" & mytablex.Fields("unidad")

    End If

    mytablex.Close
    mytablex.Open "SELECT * FROM precios where  producto='" & xproducto & "' and local='" & Combo4.Text & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        'MsgBox "Hola"
        xcosto = xcostou
        campo_precios(0).unidad = xunidad
        campo_precios(0).factor = xfactor
        campo_precios(0).precio = "" '& mytablex.Fields("costou")
        campo_precios(0).costo = xcostou
        xbuf = calcula_saldo(xsaldo, xfactor)
        campo_precios(0).stock = "" & xbuf
        xmargen = 0
        campo_precios(0).margen = "" & xmargen
        '----------------------------------------------
        xcosto = 0

        If Val("" & mytablex.Fields("factor1")) > 0 Then
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor1"))
            campo_precios(1).unidad = "" & mytablex.Fields("unidad1")
            campo_precios(1).factor = "" & mytablex.Fields("factor1")
            campo_precios(1).precio = "" & mytablex.Fields("pventa1")
            campo_precios(1).costo = "" & xcosto
            xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor1")))
            campo_precios(1).stock = "" & xbuf
            xmargen = 0

            If xcosto > 0 Then
                xmargen = ((Val("" & mytablex.Fields("pventa1")) - xcosto) * 100) / xcosto

            End If

            campo_precios(1).margen = "" & xmargen

            '--------
        End If

        '---------
        If Val("" & mytablex.Fields("factor2")) > 0 Then
            campo_precios(2).unidad = "" & mytablex.Fields("unidad2")
            campo_precios(2).factor = "" & mytablex.Fields("factor2")
            campo_precios(2).precio = "" & mytablex.Fields("pventa2")
            xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor2")))
            campo_precios(2).stock = "" & xbuf
            xcosto = 0
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor2"))
            campo_precios(2).costo = "" & xcosto
            xmargen = 0

            If xcosto > 0 Then
                xmargen = ((Val("" & mytablex.Fields("pventa2")) - xcosto) * 100) / xcosto

            End If

            campo_precios(2).margen = "" & xmargen

        End If
   
        If Val("" & mytablex.Fields("factor3")) > 0 Then
            campo_precios(3).unidad = "" & mytablex.Fields("unidad3")
            campo_precios(3).factor = "" & mytablex.Fields("factor3")
            campo_precios(3).precio = "" & mytablex.Fields("pventa3")
            xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor3")))
            campo_precios(3).stock = "" & xbuf
            xcosto = 0
   
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor3"))
   
            campo_precios(3).costo = "" & xcosto
            xmargen = 0

            If xcosto > 0 Then
                xmargen = ((Val("" & mytablex.Fields("pventa3")) - xcosto) * 100) / xcosto
                campo_precios(3).margen = "" & xmargen

            End If

            campo_precios(3).margen = "" & xmargen

        End If

        If Val("" & mytablex.Fields("factor4")) > 0 Then
            campo_precios(4).unidad = "" & mytablex.Fields("unidad4")
            campo_precios(4).factor = "" & mytablex.Fields("factor4")
            campo_precios(4).precio = "" & mytablex.Fields("pventa4")
            xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor4")))
            campo_precios(4).stock = "" & xbuf
            xcosto = 0
   
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor4"))
   
            campo_precios(4).costo = "" & xcosto
            xmargen = 0

            If xcosto > 0 Then
                xmargen = ((Val("" & mytablex.Fields("pventa4")) - xcosto) * 100) / xcosto

            End If

            campo_precios(4).margen = "" & xmargen

        End If
   
        If Val("" & mytablex.Fields("factor5")) > 0 Then
            campo_precios(5).unidad = "" & mytablex.Fields("unidad5")
            campo_precios(5).factor = "" & mytablex.Fields("factor5")
            campo_precios(5).precio = "" & mytablex.Fields("pventa5")
            xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor5")))
            campo_precios(5).stock = "" & xbuf
            xcosto = 0
   
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor5"))
   
            campo_precios(5).costo = "" & xcosto
            xmargen = 0

            If xcosto > 0 Then
                xmargen = ((Val("" & mytablex.Fields("pventa5")) - xcosto) * 100) / xcosto

            End If

            campo_precios(5).margen = "" & xmargen

        End If
   
        If Val("" & mytablex.Fields("factor6")) > 0 Then
            campo_precios(6).unidad = "" & mytablex.Fields("unidad6")
            campo_precios(6).factor = "" & mytablex.Fields("factor6")
            campo_precios(6).precio = "" & mytablex.Fields("pventa6")
            xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor6")))
            campo_precios(6).stock = "" & xbuf
            xcosto = 0
   
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor6"))
   
            campo_precios(6).costo = "" & xcosto
            xmargen = 0

            If xcosto > 0 Then
                xmargen = ((Val("" & mytablex.Fields("pventa5")) - xcosto) * 100) / xcosto
         
            End If

            campo_precios(6).margen = "" & xmargen

        End If
   
        If Val("" & mytablex.Fields("factor7")) > 0 Then
            campo_precios(7).unidad = "" & mytablex.Fields("unidad7")
            campo_precios(7).factor = "" & mytablex.Fields("factor7")
            campo_precios(7).precio = "" & mytablex.Fields("pventa7")
            xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor7")))
            campo_precios(7).stock = "" & xbuf
            xcosto = 0
   
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor7"))
   
            campo_precios(7).costo = "" & xcosto
            xmargen = 0

            If xcosto > 0 Then
                xmargen = ((Val("" & mytablex.Fields("pventa7")) - xcosto) * 100) / xcosto

            End If

            campo_precios(7).margen = "" & xmargen

        End If
   
        If Val("" & mytablex.Fields("factor8")) > 0 Then
            campo_precios(8).unidad = "" & mytablex.Fields("unidad8")
            campo_precios(8).factor = "" & mytablex.Fields("factor8")
            campo_precios(8).precio = "" & mytablex.Fields("pventa8")
            xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor8")))
            campo_precios(8).stock = "" & xbuf
            xcosto = 0
   
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor8"))
   
            campo_precios(8).costo = "" & xcosto
            xmargen = 0

            If xcosto > 0 Then
                xmargen = ((Val("" & mytablex.Fields("pventa8")) - xcosto) * 100) / xcosto

            End If

            campo_precios(8).margen = "" & xmargen

        End If
   
        If Val("" & mytablex.Fields("factor9")) > 0 Then
            campo_precios(9).unidad = "" & mytablex.Fields("unidad9")
            campo_precios(9).factor = "" & mytablex.Fields("factor9")
            campo_precios(9).precio = "" & mytablex.Fields("pventa9")
            xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor9")))
            campo_precios(9).stock = "" & xbuf
            xcosto = 0
   
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor9"))
   
            campo_precios(9).costo = "" & xcosto
            xmargen = 0

            If xcosto > 0 Then
                xmargen = ((Val("" & mytablex.Fields("pventa9")) - xcosto) * 100) / xcosto
         
            End If

            campo_precios(9).margen = "" & xmargen

        End If
   
        If Val("" & mytablex.Fields("factor10")) > 0 Then
            campo_precios(10).unidad = "" & mytablex.Fields("unidad10")
            campo_precios(10).factor = "" & mytablex.Fields("factor10")
            campo_precios(10).precio = "" & mytablex.Fields("pventa10")
            xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor10")))
            campo_precios(10).stock = "" & xbuf
            xcosto = 0
   
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor10"))
   
            campo_precios(10).costo = "" & xcosto
            xmargen = 0

            If xcosto > 0 Then
                xmargen = ((Val("" & mytablex.Fields("pventa10")) - xcosto) * 100) / xcosto

            End If

            campo_precios(10).margen = "" & xmargen

        End If

        'margenes
        sw = 1
   
    End If

    mytablex.Close

    dbgrid4.refresh
    Frame5.Visible = True
    dbgrid4.SetFocus
    Exit Sub
cmd89012_err:
    MsgBox "Error en carga Grid " + error$, 48, "Aviso"
    Exit Sub

End Sub

Function busca_tipox(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    'Label16 = ""
    'acu1 = "0"
    mytablex.Open "SELECT * FROM tipo where   tipo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        busca_tipox = "" & mytablex.Fields("tipodoc")

    End If

    mytablex.Close

End Function

Function valida_flag(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM tipo where  tipo='" & ttipo & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
       
        Select Case "" & mytablex.Fields("tipodoc")

            Case "Z"
                valida_flag = 3

            Case "1", "T", "A", "B", "C", "D", "G", "E", "F" 'VENTAS
                valida_flag = 1
       
                '12/06/2017 kenyo COSTO Y NO PV EN ORDEN DE COMPRA
                ' Case "S", "J", "K", "L", "M", "P", "N", "O" 'COMPRAS
            Case "S", "J", "K", "L", "M", "P", "N", "O", "R" 'COMPRAS. SI ES R SALDRA SOLO COSTO EN ORDEN DE COMPRA
                '12/06/2017 kenyo COSTO Y NO PV EN ORDEN DE COMPRA
       
                valida_flag = 2

        End Select

    End If

    mytablex.Close

End Function

Function graba_adelantos(buf1 As String, _
                         buf2 As String, _
                         buf3 As String, _
                         buf4 As String, _
                         xsw As String)

    Dim mytablex As New ADODB.Recordset

    If Len(buf1) = 0 Then Exit Function
    If Len(buf2) = 0 Then Exit Function

    mytablex.Open "SELECT * FROM recibo where  local='" & buf1 & "' and tipo='" & buf2 & "' and serie='" & buf3 & "' and numero='" & buf4 & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        'mytablex.Edit
        mytablex.Fields("usado") = xsw
        mytablex.Update
        graba_adelantos = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Sub graba_yausado_guia(xsw As String)

    'MsgBox cgusuario & " " & acu1
    If cgusuario <> "FACTURA" Then Exit Sub 'verificamos si es guia o factura
    If acu1 <> "S" And acu1 <> "T" Then Exit Sub
    descarga_el_uso local1, "" & tipo1, "" & serie1, "" & numero1, xsw
    descarga_el_uso local1, "" & tipo1, "" & serie2, "" & numero2, xsw
    descarga_el_uso local1, "" & tipo1, "" & serie3, "" & numero3, xsw
    descarga_el_uso local1, "" & tipo1, "" & serie4, "" & numero4, xsw
    descarga_el_uso local1, "" & tipo1, "" & serie5, "" & numero5, xsw
    descarga_el_uso local1, "" & tipo1, "" & serie6, "" & numero6, xsw
    descarga_el_uso local1, "" & tipo1, "" & serie7, "" & numero7, xsw

End Sub

Sub descarga_el_uso(buf0 As String, _
                    buf1 As String, _
                    buf2 As String, _
                    buf3 As String, _
                    xsw As String)

    On Error GoTo cmd8912d

    Dim mytablex As New ADODB.Recordset

    If Len(buf0) = 0 Then Exit Sub
    If Len(buf1) = 0 Then Exit Sub
    If Len(buf2) = 0 Then Exit Sub
    If Len(buf3) = 0 Then Exit Sub

    mytablex.Open "SELECT * FROM " & cgusuario & " where  local='" & buf0 & "' and tipo='" & buf1 & "' and serie='" & buf2 & "' and numero='" & buf3 & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        mytablex.Fields("yausado") = xsw
        mytablex.Update

    End If

    '------------------------------------- ------------
    mytablex.Close
    Exit Sub
cmd8912d:
    MsgBox "Aviso en descarga el uso " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub consulta_facturacion_anula()
    cerrar_data1
    sw_consulta = 0
    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "Tipo"
    Combo2.AddItem "Serie"
    Combo2.AddItem "Numero"
    Combo2.AddItem "Codigo"
    Combo2.AddItem "Nombre"
    Combo2.AddItem "Moneda"
    Combo2.ListIndex = 0

    Combo1.Clear
    Combo1.AddItem "Tipo"
    Combo1.AddItem "Serie"
    Combo1.AddItem "Numero"
    Combo1.AddItem "Codigo"
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Moneda"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "30"
    Command1_Click

End Sub

Function crea_nuevos_proveedores(buf1 As String, _
                                 buf2 As String, _
                                 buf3 As String, _
                                 buf4 As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM codprov where  codigo='" & buf1 & "' and producto='" & buf2 & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        'mytablex.Edit
        mytablex.Fields("codigo") = "" & buf1
        mytablex.Fields("producto") = "" & buf2
        mytablex.Fields("costo") = Val("" & buf3)

        If Len(buf4) = 10 Then
            mytablex.Fields("fecha") = Format(buf4, "dd/mm/yyyy")

        End If

        mytablex.Update
    Else
        mytablex.AddNew
        mytablex.Fields("codigo") = "" & buf1
        mytablex.Fields("producto") = "" & buf2
        mytablex.Fields("costo") = Val("" & buf3)

        If Len(buf4) = 10 Then
            mytablex.Fields("fecha") = Format(buf4, "dd/mm/yyyy")

        End If

        mytablex.Update

    End If

    mytablex.Close

End Function

Function crea_nuevos_clientes(buf1 As String, _
                              buf2 As String, _
                              buf3 As String, _
                              buf4 As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM codclie where  codigo='" & buf1 & "' and producto='" & buf2 & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        'mytablex.Edit
        mytablex.Fields("codigo") = "" & buf1
        mytablex.Fields("producto") = "" & buf2
        mytablex.Fields("costo") = Val("" & buf3)

        If Len(buf4) = 10 Then
            mytablex.Fields("fecha") = Format(buf4, "dd/mm/yyyy")

        End If

        mytablex.Update
    Else
        mytablex.AddNew
        mytablex.Fields("codigo") = "" & buf1
        mytablex.Fields("producto") = "" & buf2
        mytablex.Fields("costo") = Val("" & buf3)

        If Len(buf4) = 10 Then
            mytablex.Fields("fecha") = Format(buf4, "dd/mm/yyyy")

        End If

        mytablex.Update

    End If

    mytablex.Close

End Function

Function busca_cod_prov(buf1 As String, buf2 As String)

    Dim mytablex As New ADODB.Recordset

    If Len(codigo) = 0 Then Exit Function
    mytablex.Open "SELECT * FROM codprov where  codigo='" & codigo & "' and codigoP='" & buf2 & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        buf2 = "" & mytablex.Fields("producto")
        busca_cod_prov = 1

    End If

    mytablex.Close

End Function

Function busca_equiva(buf As String) As Integer

    Dim buf1     As String

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Integer

    Dim I        As Integer

    buf1 = ""

    If flag_denisse = "1" Then
        sdx = 18 - Len(buf)

        For I = 1 To sdx
            buf1 = buf1 & "0"
        Next I

    End If

    buf1 = buf1 & buf

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM productb where barras='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        buf = "" & mytablex.Fields("producto")
        busca_equiva = 1
        mytablex.Close
        Exit Function

    End If

    mytablex.Close
   
    mytablex.Open "SELECT * FROM producto where barras='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        buf = "" & mytablex.Fields("producto")
        busca_equiva = 1

    End If

    mytablex.Close

End Function

Function busca_caja()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM parameca where  caja='" & caja & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        busca_caja = 1

    End If

    mytablex.Close

End Function

Function busca_turno()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM turno where  turno='" & turno & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        busca_turno = 1

    End If

    mytablex.Close

End Function

Function salir_sin_grabar()

    Dim found As Integer

    If Frame7.Visible = True Then Exit Function
    If Frame4.Visible = True Then Exit Function
    If dbgrid3.Visible = True Then Exit Function
    If Frame2.Visible = True Then Exit Function
    If Frame1.Visible = True Then Exit Function
    If MsgBox("Desea Salir Grabando Lo digitado?", 1, "Aviso") <> 1 Then Exit Function
    If Len(codigo) = 0 Or Len(serie) = 0 Or Len(Numero) = 0 Then ' si es datos principales sin datos solo salir
        salir_sin_grabar = 1
        Exit Function

    End If

    found = valida()

    If found = 0 Then
        MsgBox "Campos Invalidos", 48, "Aviso"
        Exit Function

    End If

    sumar_detalle
    found = grabar1()

    If found = 0 Then
        MsgBox "No se pudo grabar ", 48, "Aviso"
        Exit Function

    End If

    salir_sin_grabar = 1

End Function

Function grabar1()

    Dim rs        As Recordset

    Dim I         As Integer

    Dim pracu     As String

    Dim buf1      As String

    Dim found     As Integer

    Dim mytablex  As New ADODB.Recordset

    Dim mytablexy As New ADODB.Recordset

    Dim te        As String

    Dim ts        As String

    Dim xc1       As Double

    Dim xc2       As Double

    Dim xc3       As Double

    Dim xc4       As Double

    Dim fila      As Integer

    Dim sw        As Integer

    sw = 0
    'Set mytablexy = mydbxglo.OpenTable(dgusuariog)
    'mytablexy.Index = "tdetalle"

    mytablex.Open "SELECT * FROM " & cgusuario & " where  local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & Numero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then  'si existe
        mytablex.AddNew
        grabando mytablex
        mytablex.Fields("estado") = "0"
        mytablex.Fields("yausado") = "0"
        mytablex.Update
        grabar1 = 1
    Else
        'mytablex.Edit
        grabando mytablex
        mytablex.Fields("estado") = "0"
        mytablex.Fields("yausado") = "0"
        mytablex.Update
        grabar1 = 1

    End If

    mytablex.Close

    cn.Execute ("delete from " & dgusuariog & " where local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & Numero & "'")
    mytablexy.Open "SELECT * FROM " & dgusuariog & " where local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & Numero & "'", cn, adOpenKeyset, adLockOptimistic
      
    Data2.refresh
    Set rs = Data2.Recordset.Clone
    Do

        If rs.EOF Then Exit Do
        mytablexy.AddNew

        For I = 0 To rs.Fields.count - 2
            mytablexy.Fields(I) = rs.Fields(I)
        Next I

        mytablexy.Fields("local") = "" & local1
        mytablexy.Fields("tipo") = "" & ttipo
        mytablexy.Fields("serie") = "" & serie
        mytablexy.Fields("numero") = "" & Numero
        mytablexy.Fields("vendedor") = "" & vendedor
        mytablexy.Fields("moneda") = "" & moneda
        mytablexy.Fields("bodega") = "" & bodega
        mytablexy.Fields("codigo") = "" & codigo
        mytablexy.Fields("localf") = "" & localf
        mytablexy.Fields("bodegaf") = "" & bodegaf
        mytablexy.Fields("acu") = "" & racu
        mytablexy.Fields("acu1") = ""
        'mytablexy.Fields("acu1") = "" & acu1
        mytablexy.Fields("flage") = "" & flage
        mytablexy.Fields("tipoclie") = tipoclie
        mytablexy.Fields("usuario") = "" & gusuario

        If Len(cajero) > 0 Then
            mytablexy.Fields("usuario") = "" & cajero

        End If

        mytablexy.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
        'mytablexy.Fields("hora") = Format(hora, "hh:MM")
        mytablexy.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
        mytablexy.Fields("estado") = "0"
        mytablexy.Fields("caja") = caja

        If Len(caja) = 0 Then
            mytablexy.Fields("caja") = "00"

        End If

        mytablexy.Fields("turno") = turno
        mytablexy.Fields("servicio") = Servicio
        mytablexy.Update
        grabar1 = 1
        rs.MoveNext
    Loop
    mytablexy.Close

End Function

Function ver_cambio_precios(buf As String)

End Function

Function ver_docena1(mytablex As Table)

    Dim xbuf1(10) As String

    Dim xbuf2(10) As Double

    Dim xbuf3(10) As Double

    Dim j         As Integer

    Dim I         As Integer

    Dim sdx       As Double

    xbuf1(1) = "" & mytablex.Fields("unidad1")
    xbuf1(1) = "" & mytablex.Fields("unidad2")
    xbuf1(2) = "" & mytablex.Fields("unidad3")
    xbuf1(3) = "" & mytablex.Fields("unidad4")
    xbuf1(4) = "" & mytablex.Fields("unidad5")
    xbuf1(5) = "" & mytablex.Fields("unidad6")
    xbuf1(6) = "" & mytablex.Fields("unidad7")
    xbuf1(7) = "" & mytablex.Fields("unidad8")
    xbuf1(8) = "" & mytablex.Fields("unidad9")
    xbuf1(9) = "" & mytablex.Fields("unidad10")

    xbuf2(0) = Val("" & mytablex.Fields("factor1"))
    xbuf2(1) = Val("" & mytablex.Fields("factor2"))
    xbuf2(2) = Val("" & mytablex.Fields("factor3"))
    xbuf2(3) = Val("" & mytablex.Fields("factor4"))
    xbuf2(4) = Val("" & mytablex.Fields("factor5"))
    xbuf2(5) = Val("" & mytablex.Fields("factor6"))
    xbuf2(6) = Val("" & mytablex.Fields("factor7"))
    xbuf2(7) = Val("" & mytablex.Fields("factor8"))
    xbuf2(8) = Val("" & mytablex.Fields("factor9"))
    xbuf2(9) = Val("" & mytablex.Fields("factor10"))

    xbuf3(0) = Val("" & mytablex.Fields("pventa1"))
    xbuf3(1) = Val("" & mytablex.Fields("pventa2"))
    xbuf3(2) = Val("" & mytablex.Fields("pventa3"))
    xbuf3(3) = Val("" & mytablex.Fields("pventa4"))
    xbuf3(4) = Val("" & mytablex.Fields("pventa5"))
    xbuf3(5) = Val("" & mytablex.Fields("pventa6"))
    xbuf3(6) = Val("" & mytablex.Fields("pventa7"))
    xbuf3(7) = Val("" & mytablex.Fields("pventa8"))
    xbuf3(8) = Val("" & mytablex.Fields("pventa9"))
    xbuf3(9) = Val("" & mytablex.Fields("pventa10"))

    sdx = 0
    j = 0

    For I = 0 To 9

        If I = 0 Then
            sdx = xbuf2(I)
            j = I

        End If

        If xbuf2(I) > sdx Then
            sdx = xbuf2(I)
            j = I

        End If

    Next I

    If sdx > 1 Then
        dbgrid2.columns("unidad") = xbuf1(j)
        dbgrid2.columns("factor") = xbuf2(j)
        dbgrid2.columns("precio") = xbuf3(j)
        dbgrid2.columns("total") = xbuf3(j)
        dbgrid2.columns("subtotal") = xbuf3(j)

    End If

    If sdx = 0 Then  'no pasa nada

    End If

End Function

Function busca_cajero()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM vendedor where  codigo='" & cajero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        busca_cajero = 1

    End If

    mytablex.Close

End Function

Function leer_archivo_texto()

    Dim buf As String

    If Dir$(globaldir & "\fecha.TXT") <> "" Then
        Open globaldir & "\fecha.TXT" For Input As #1
        Input #1, buf
        Close #1
        fecha = buf
        fechae = buf

    End If

End Function

Function guardar_fecha()

    If Not IsDate(fecha) Then Exit Function
    Open globaldir & "\fecha.TXT" For Output As #1
    Print #1, fecha;
    Close #1
   
End Function

Function ve_descarga(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM tipo where  tipo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe

        Select Case "" & mytablex.Fields("tipodoc")

            Case "1", "A", "B", "C", "D", "G", "E", "F", "J", "K", "L", "M", "P", "N", "O", "S", "T"
                ve_descarga = 1

        End Select

    End If

    mytablex.Close

End Function

Sub ver_presenta()

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    buf = "" & dbgrid2.columns("producto")
    presenta = ""
    precio = ""
    mytablex.Open "SELECT * FROM producto where  producto='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        presenta = "" & mytablex.Fields("presenta")

    End If

    mytablex.Close
    mytablex.Open "SELECT * FROM precios where  producto='" & buf & "' and local='" & local1 & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        precio = "" & mytablex.Fields("pventa1")

    End If

    mytablex.Close

End Sub

Function tipo_costo(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM tipo where  tipo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        tipo_costo = "" & mytablex.Fields("tipocosto")

    End If

    mytablex.Close

End Function

Sub actualizar_precios(mytablex As ADODB.Recordset)

    Dim sw As Integer

    On Error GoTo cmd89121_err

    Dim mytableyy As New ADODB.Recordset

    Dim mytabley  As New ADODB.Recordset

    sw = 0

    mytableyy.Open "SELECT * FROM familia where  familia='" & "" & mytablex.Fields("familia") & "'", cn, adOpenKeyset, adLockOptimistic

    If mytableyy.RecordCount = 0 Then
        mytableyy.Close
        Exit Sub

    End If
       
    If "" & mytableyy.Fields("obliga") = "S" Then
        mytabley.Open "SELECT * FROM precios where  producto='" & "" & mytablex.Fields("producto") & "' and local='" & local1 & "'", cn, adOpenKeyset, adLockOptimistic

        If mytabley.RecordCount = 0 Then
            mytabley.AddNew
            inicializa_precios mytabley
            pone_margenes mytableyy, mytabley
            mytabley.Fields("producto") = "" & mytablex.Fields("producto")
            mytabley.Fields("local") = "" & local1
            calcula_margenes mytabley, mytablex
            mytabley.Update
        Else
            pone_margenes mytableyy, mytabley
            calcula_margenes mytabley, mytablex
            mytabley.Update

        End If

        mytabley.Close

    End If

    mytableyy.Close
      
    Exit Sub
cmd89121_err:
    MsgBox "Aviso en actualiza precios " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub calcula_margenes(mytablex As ADODB.Recordset, mytabley As ADODB.Recordset)

    Dim sdx     As Double

    Dim sdx1    As Double

    Dim sdx2    As Double

    Dim acostou As String

    On Error GoTo cmd786_err

    If Val(paridad) <= 0 Then
        paridad = "1"

    End If
      
    sdx = Val(Format(Val("" & mytabley.Fields("costou")), "0.00")) + Val("" & mytabley.Fields("flete"))
    acostou = "" & sdx

    If mytabley.Fields("monedac") = "S" Then
        If mytabley.Fields("monedav") = "D" Then
            sdx = Val(acostou) / Val(paridad)

            If sdx <= 0 Then
                sdx = 1

            End If

            acostou = "" & sdx

        End If

    End If

    If mytabley.Fields("monedac") = "D" Then
        If mytabley.Fields("monedav") = "S" Then
            sdx = Val(acostou) * Val(paridad)

            If sdx <= 0 Then
                sdx = 1

            End If

            acostou = "" & sdx

        End If

    End If

    If Val(mytablex.Fields("margen1")) > 0 And Val(mytablex.Fields("factor1")) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(mytablex.Fields("margen1")) / 100
        sdx = sdx * Val(mytablex.Fields("factor1"))
        mytablex.Fields("pventa1") = Format(sdx, "0.00")

    End If

    If Val(mytablex.Fields("margen2")) > 0 And Val(mytablex.Fields("factor2")) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(mytablex.Fields("margen2")) / 100
        sdx = sdx * Val(mytablex.Fields("factor2"))
        mytablex.Fields("pventa2") = Format(sdx, "0.00")

    End If
          
    If Val(mytablex.Fields("margen3")) > 0 And Val(mytablex.Fields("factor3")) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(mytablex.Fields("margen3")) / 100
        sdx = sdx * Val(mytablex.Fields("factor3"))
        mytablex.Fields("pventa3") = Format(sdx, "0.00")

    End If

    If Val(mytablex.Fields("margen4")) > 0 And Val(mytablex.Fields("factor4")) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(mytablex.Fields("margen4")) / 100
        sdx = sdx * Val(mytablex.Fields("factor4"))
        mytablex.Fields("pventa4") = Format(sdx, "0.00")

    End If
          
    If Val(mytablex.Fields("margen5")) > 0 And Val(mytablex.Fields("factor5")) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(mytablex.Fields("margen5")) / 100
        sdx = sdx * Val(mytablex.Fields("factor5"))
        mytablex.Fields("pventa5") = Format(sdx, "0.00")

    End If

    If Val(mytablex.Fields("margen6")) > 0 And Val(mytablex.Fields("factor6")) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(mytablex.Fields("margen6")) / 100
        sdx = sdx * Val(mytablex.Fields("factor6"))
        mytablex.Fields("pventa6") = Format(sdx, "0.00")

    End If
          
    If Val(mytablex.Fields("margen7")) > 0 And Val(mytablex.Fields("factor7")) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(mytablex.Fields("margen7")) / 100
        sdx = sdx * Val(mytablex.Fields("factor7"))
        mytablex.Fields("pventa7") = Format(sdx, "0.00")

    End If

    If Val(mytablex.Fields("margen8")) > 0 And Val(mytablex.Fields("factor8")) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(mytablex.Fields("margen8")) / 100
        sdx = sdx * Val(mytablex.Fields("factor8"))
        mytablex.Fields("pventa8") = Format(sdx, "0.00")

    End If
          
    If Val(mytablex.Fields("margen9")) > 0 And Val(mytablex.Fields("factor9")) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(mytablex.Fields("margen9")) / 100
        sdx = sdx * Val(mytablex.Fields("factor9"))
        mytablex.Fields("pventa9") = Format(sdx, "0.00")

    End If

    If Val(mytablex.Fields("margen10")) > 0 And Val(mytablex.Fields("factor10")) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(mytablex.Fields("margen10")) / 100
        sdx = sdx * Val(mytablex.Fields("factor10"))
        mytablex.Fields("pventa10") = Format(sdx, "0.00")

    End If
          
    If Val(mytablex.Fields("margen11")) > 0 And Val(mytablex.Fields("factor1")) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(mytablex.Fields("margen11")) / 100
        sdx = sdx * Val(mytablex.Fields("factor1"))
        mytablex.Fields("pventa11") = Format(sdx, "0.00")

    End If
          
    If Val(mytablex.Fields("margen12")) > 0 And Val(mytablex.Fields("factor1")) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(mytablex.Fields("margen12")) / 100
        sdx = sdx * Val(mytablex.Fields("factor1"))
        mytablex.Fields("pventa12") = Format(sdx, "0.00")

    End If
          
    If Val(mytablex.Fields("margen13")) > 0 And Val(mytablex.Fields("factor1")) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(mytablex.Fields("margen13")) / 100
        sdx = sdx * Val(mytablex.Fields("factor1"))
        mytablex.Fields("pventa13") = Format(sdx, "0.00")

    End If
          
    If Val(mytablex.Fields("margen14")) > 0 And Val(mytablex.Fields("factor1")) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(mytablex.Fields("margen14")) / 100
        sdx = sdx * Val(mytablex.Fields("factor1"))
        mytablex.Fields("pventa14") = Format(sdx, "0.00")

    End If
          
    If Val(mytablex.Fields("margen15")) > 0 And Val(mytablex.Fields("factor1")) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(mytablex.Fields("margen15")) / 100
        sdx = sdx * Val(mytablex.Fields("factor1"))
        mytablex.Fields("pventa15") = Format(sdx, "0.00")

    End If

    Exit Sub
cmd786_err:
    MsgBox "Error en calcula margenes", 48, "Aviso"
    Exit Sub

End Sub

Sub pone_margenes(mytablex As ADODB.Recordset, mytabley As ADODB.Recordset)

    If Val("" & mytablex.Fields("margen1")) > 0 And Val("" & mytabley.Fields("factor1")) > 0 Then
        mytabley.Fields("margen1") = "" & mytablex.Fields("margen1")

    End If

    If Val("" & mytablex.Fields("margen2")) > 0 And Val("" & mytabley.Fields("factor2")) > 0 Then
        mytabley.Fields("margen2") = "" & mytablex.Fields("margen2")

    End If

    If Val("" & mytablex.Fields("margen3")) > 0 And Val("" & mytabley.Fields("factor3")) > 0 Then
        mytabley.Fields("margen3") = "" & mytablex.Fields("margen3")

    End If

    If Val("" & mytablex.Fields("margen4")) > 0 And Val("" & mytabley.Fields("factor4")) > 0 Then
        mytabley.Fields("margen4") = "" & mytablex.Fields("margen4")

    End If

    If Val("" & mytablex.Fields("margen5")) > 0 And Val("" & mytabley.Fields("factor5")) > 0 Then
        mytabley.Fields("margen5") = "" & mytablex.Fields("margen5")

    End If

    If Val("" & mytablex.Fields("margen6")) > 0 And Val("" & mytabley.Fields("factor6")) > 0 Then
        mytabley.Fields("margen6") = "" & mytablex.Fields("margen6")

    End If

    If Val("" & mytablex.Fields("margen7")) > 0 And Val("" & mytabley.Fields("factor7")) > 0 Then
        mytabley.Fields("margen7") = "" & mytablex.Fields("margen7")

    End If

    If Val("" & mytablex.Fields("margen8")) > 0 And Val("" & mytabley.Fields("factor8")) > 0 Then
        mytabley.Fields("margen8") = "" & mytablex.Fields("margen8")

    End If

    If Val("" & mytablex.Fields("margen9")) > 0 And Val("" & mytabley.Fields("factor9")) > 0 Then
        mytabley.Fields("margen9") = "" & mytablex.Fields("margen9")

    End If

    If Val("" & mytablex.Fields("margen10")) > 0 And Val("" & mytabley.Fields("factor10")) > 0 Then
        mytabley.Fields("margen10") = "" & mytablex.Fields("margen10")

    End If

End Sub

Sub inicializa_precios(mytablex As ADODB.Recordset)
    mytablex.Fields("pm1") = 0
    mytablex.Fields("pm2") = 0
    mytablex.Fields("pm3") = 0
    mytablex.Fields("pm4") = 0
    mytablex.Fields("pm5") = 0
    mytablex.Fields("pm6") = 0
    mytablex.Fields("pm7") = 0
    mytablex.Fields("pm8") = 0
    mytablex.Fields("pm9") = 0
    mytablex.Fields("pm10") = 0

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

End Sub

Sub refresca_precios()

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd700_err

    dbgrid89.Visible = False

    If opcion1 = "8" Or opcion1 = "888" Or opcion1 = "50" Then
        dbgrid89.Visible = True
        Set mytablex = Nothing
        mytablex.Open "select Unidad1,Pventa1,Unidad2,Pventa2,Unidad3,Pventa3,Unidad4,Pventa4 from Precios where producto='" & "" & DBGrid1.columns("producto") & "' and local='01'", cn, adOpenStatic, adLockOptimistic
        Set dbgrid89.DataSource = mytablex
        dbgrid89.columns(0).Width = 1000
        dbgrid89.columns(1).Width = 1000
        dbgrid89.columns(2).Width = 1000
        dbgrid89.columns(3).Width = 1000
        dbgrid89.columns(4).Width = 1000
        dbgrid89.columns(5).Width = 1000

    End If

    Exit Sub
cmd700_err:
    Exit Sub

End Sub

Sub actualiza_receta(buf As String, bufc As String)

    On Error GoTo cmd9093_err

    cn.Execute ("update receta set precio=" & Val(bufc) & ",total=cantidad*" & Val(bufc) & " where productoi='" & buf & "'")

    Exit Sub
cmd9093_err:
    Exit Sub

End Sub

Public Function rRedondear(ByVal cantidad As Currency, _
                           Optional redondeo As Byte = 2) As Currency

    Dim dblPot As Double

    Dim dblF   As Double

    If cantidad < 0 Then dblF = -0.5 Else dblF = 0.5
    dblPot = 10 ^ redondeo
    rRedondear = Fix(cantidad * dblPot * (1 + 1E-16) + dblF) / dblPot

End Function

Private Function rRedondeo(ByVal Numero, ByVal Decimales)
    rRedondeo = Int(Numero * 10 ^ Decimales + 1 / 2) / 10 ^ Decimales

End Function

Sub resuma_precios(xpercepcion As Double)

    On Error GoTo cmd94534_err

    Data2.Recordset.Fields("neto") = Val("" & Data2.Recordset.Fields("cantidad")) * Val("" & Data2.Recordset.Fields("precio"))
    Data2.Recordset.Fields("descuento") = Val("" & Data2.Recordset.Fields("neto")) * Val("" & Data2.Recordset.Fields("deslipo")) / 100 + Val("" & Data2.Recordset.Fields("neto")) * Val("" & Data2.Recordset.Fields("destopo")) / 100
    Data2.Recordset.Fields("total") = Val("" & Data2.Recordset.Fields("neto")) - Val("" & Data2.Recordset.Fields("descuento")) 'cobrar
    Data2.Recordset.Fields("subtotal") = Val("" & Data2.Recordset.Fields("total")) / (1 + Val("" & Data2.Recordset.Fields("igv")) / 100) 'calcular descuento
    Data2.Recordset.Fields("impuesto") = Val("" & Data2.Recordset.Fields("total")) - Val("" & Data2.Recordset.Fields("subtotal")) 'cobrar
    Data2.Recordset.Fields("tivap") = Val("" & Data2.Recordset.Fields("total")) * Val("" & Data2.Recordset.Fields("ivap")) / 100
    Data2.Recordset.Fields("tdetra") = Val("" & Data2.Recordset.Fields("total")) * Val("" & Data2.Recordset.Fields("detraccion")) / 100   'calcular descuento
    Data2.Recordset.Fields("tpercepcio") = Val("" & Data2.Recordset.Fields("total")) * Val("" & Data2.Recordset.Fields("percepcion")) / 100   'calcular descuento

    If Trim(menup.Label10) = "ARGENTINA" Then
        Data2.Recordset.Fields("tpercepcio") = Val("" & Data2.Recordset.Fields("subtotal")) * Val("" & Data2.Recordset.Fields("percepcion")) / 100   'calcular descuento

    End If

    Data2.Recordset.Fields("total") = Val("" & Data2.Recordset.Fields("total")) + Val("" & Data2.Recordset.Fields("tpercepcio")) - Val("" & Data2.Recordset.Fields("tdetra")) 'cobrar
    Data2.Recordset.Fields("servicioco") = Val("" & Data2.Recordset.Fields("subtotal")) * Val("" & Data2.Recordset.Fields("serviciopo")) / 100      'calcular descuento
    Data2.Recordset.Fields("tisc") = Val("" & Data2.Recordset.Fields("subtotal")) * Val("" & Data2.Recordset.Fields("isc")) / 100
    Data2.Recordset.Fields("total") = rRedondear(Val("" & Data2.Recordset.Fields("total")), 2)
    Exit Sub
cmd94534_err:
    MsgBox "Aviso en resuma_precios ", 48, "Aviso"
    Exit Sub

End Sub

Sub graba_temporaldsi(mytablex As ADODB.Recordset, _
                      mytabley As ADODB.Recordset, _
                      sw As Integer, _
                      canti As Double)

    Dim found    As Integer

    Dim pventa1  As Double

    Dim costou   As Double

    Dim buf      As String

    Dim mytables As New ADODB.Recordset

    pventa1 = Val("" & mytabley.Fields("pventa1"))
    costou = Val(Format(Val("" & mytablex.Fields("costou")), "0.00"))

    If Val(paridad) <= 0 Then
        paridad = "1"

    End If

    'MsgBox moneda
    If "" & moneda = "S" Then
        If "" & mytablex.Fields("monedav") = "D" Then
            pventa1 = Val("" & mytabley.Fields("pventa1")) * Val(paridad)

        End If

        If "" & mytablex.Fields("monedaC") = "D" Then
            costou = costou * Val(paridad)

        End If

    End If

    If "" & moneda = "D" Then
        If "" & mytablex.Fields("monedav") = "S" Then
            pventa1 = Val("" & mytabley.Fields("pventa1")) / Val(paridad)

        End If

        If "" & mytablex.Fields("monedaC") = "S" Then
            costou = costou / Val(paridad)

        End If

    End If

    mytables.Open "SELECT * FROM DUENO where  local='" & local1 & "' and producto='" & "" & mytablex.Fields("producto") & "' ", cn, adOpenKeyset, adLockOptimistic

    If mytables.RecordCount > 0 Then  'si existe
        Data2.Recordset.Fields("ccosto") = Trim("" & mytables.Fields("codigo"))

    End If

    mytables.Close
    Data2.Recordset.Fields("proveedorp") = Format(costou, "0.00") 'costo anterior
    Data2.Recordset.Fields("producto") = "" & mytablex.Fields("producto")
    'data2.recordset.fields("proveedorp") = "" '& mytablex.Fields("proveedor1")
    Data2.Recordset.Fields("tipo") = "" & ttipo
    Data2.Recordset.Fields("serie") = "" & serie
    Data2.Recordset.Fields("numero") = "" & Numero
    Data2.Recordset.Fields("vendedor") = "" & vendedor
    Data2.Recordset.Fields("descripcio") = Mid$("" & mytablex.Fields("descripcio"), 1, 60)
    Data2.Recordset.Fields("cantidad") = 1

    If canti > 0 Then
        Data2.Recordset.Fields("cantidad") = canti

    End If

    Data2.Recordset.Fields("unidad") = "" & mytabley.Fields("unidad1")
    Data2.Recordset.Fields("factor") = Val("" & mytabley.Fields("factor1"))
    Data2.Recordset.Fields("precio") = pventa1
    Data2.Recordset.Fields("total") = pventa1
    Data2.Recordset.Fields("subtotal") = pventa1
    Data2.Recordset.Fields("descuento") = 0
    Data2.Recordset.Fields("isc") = Val("" & mytablex.Fields("isc"))

    'data2.recordset.fields(13) = Val("" & mytablex.Fields("tax"))
    If valida_flag("" & racu) = "2" Then  'compras
        Data2.Recordset.Fields("unidad") = "" & mytablex.Fields("unidad")
        Data2.Recordset.Fields("factor") = Val("" & mytablex.Fields("factor"))
        Data2.Recordset.Fields("precio") = costou * Val("" & mytablex.Fields("factor"))
        Data2.Recordset.Fields("total") = costou * Val("" & mytablex.Fields("factor"))
        Data2.Recordset.Fields("subtotal") = costou * Val("" & mytablex.Fields("factor"))

    End If

    If valida_flag("" & racu) = "1" Then 'ventas
        Data2.Recordset.Fields("unidad") = "" & mytabley.Fields("unidad1")
        Data2.Recordset.Fields("factor") = Val("" & mytabley.Fields("factor1"))
        Data2.Recordset.Fields("precio") = pventa1
        Data2.Recordset.Fields("total") = pventa1
        Data2.Recordset.Fields("subtotal") = pventa1

    End If

    buf = tipo_costo("" & ttipo)

    Select Case buf

        Case "V"
            Data2.Recordset.Fields("precio") = pventa1

    End Select

    Data2.Recordset.Fields("deslipo") = 0
    Data2.Recordset.Fields("tax") = 0
    Data2.Recordset.Fields("flete") = Val("" & mytablex.Fields("flete"))
    Data2.Recordset.Fields("impuesto") = 0
    Data2.Recordset.Fields("ivap") = Val("" & mytablex.Fields("ivap"))
    Data2.Recordset.Fields("igv") = Val("" & mytablex.Fields("igv"))
    Data2.Recordset.Fields("percepcion") = Val("" & mytablex.Fields("percepcion"))
    Data2.Recordset.Fields("linea") = "" & mytablex.Fields("linea")

    Data2.Recordset.Fields("descuento") = 0
    Data2.Recordset.Fields("neto") = 0

    '---------pone a quien pertenece --------------------
    Data2.Recordset.Fields("l1") = "" '& mytablex.Fields("c11")
    Data2.Recordset.Fields("l2") = "" '& mytablex.Fields("c12")
    Data2.Recordset.Fields("l3") = "" '& mytablex.Fields("c13")
    Data2.Recordset.Fields("l4") = "" '& mytablex.Fields("c14")

    'LAS FAMILIAS+SUBFAMILIA+MARCA+SECCION
    Data2.Recordset.Fields("familia") = "" & mytablex.Fields("Familia")
    Data2.Recordset.Fields("subfamilia") = "" & mytablex.Fields("subFamilia")
    Data2.Recordset.Fields("marca") = "" & mytablex.Fields("marca")
    'data2.recordset.fields("hora") = Format(hora, "hh:MM:ss")
    Data2.Recordset.Fields("hora") = Format(Now, "hh:MM:ss")

    'If bodega = "01" Then
    '   found = ver_docena1(mytabley)
    'End If
    If racu = "E" Or racu = "N" Then  'si es nota credito por compras o ventas
        If Val("" & Data2.Recordset.Fields("precio")) >= 0 Then
            Data2.Recordset.Fields("precio") = -Val("" & Data2.Recordset.Fields("precio"))

        End If

    End If

    '-----------------------------
    'calcula_igv 0
End Sub

Function busca_productosi(buf As String, sw As Integer, canti As Double)

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim xbuf     As String

    Dim found    As Integer

    Dim sw1      As Integer

    Dim ybuf     As String

    Dim buf1     As String

    Dim I        As Integer

    Dim ssw      As Integer

    Dim sfound   As String

    If Val(paridad) <= 0 Then
        paridad = "1"

    End If

    xbuf = buf
    sw1 = 0
    ybuf = ""
    sfound = "V"

    'If valida_flag("" & racu) = 2 Then    'compras
    If Mid$(buf, 1, 1) = "!" Then   'si es codigo de proveedor
        xbuf = Mid$(buf, 2, Len(buf) - 1)

        If Len(xbuf) = 0 Then
            Exit Function

        End If

        ybuf = xbuf
   
        found = busca_cod_prov("" & codigo, xbuf)

        If found = 0 Then
            Exit Function

        End If

        found = verifica_doble("" & xbuf)

        If found = 1 Then
            Exit Function

        End If

    End If

    sw = 0
    'validamos si es que tiene busqueda por codigo proveedor
    buf1 = xbuf

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM producto where producto='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        found = busca_equiva(buf1) 'busca en la table codigo barras

        If found = 0 Then
            Exit Function

        End If

        mytablex.Open "SELECT * FROM producto where producto='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.Close
            Exit Function

        End If

    End If

    'miramos si es compra
    'sfound = tipo_costo("" & ttipo)
    'If Trim(sfound) = "C" Or Trim(sfound) = "P" Or Len(Trim(sfound)) = 0 Then
    '   graba_temporald mytablex, mytabley, sw, sfound
    '   sw1 = 1
    '   busca_producto = 1
    '   mytablex.Close
    '   mytabley.Close
    '   Exit Function
    ' End If
    '-- ahora busca los precios
a134:
    mytabley.Open "SELECT * FROM precios where  producto='" & "" & mytablex.Fields("producto") & "' and local='" & local1 & "'", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount = 0 Then  'si existe
        mytabley.AddNew
        mytabley.Fields("local") = Trim("" & local1)
        mytabley.Fields("producto") = Trim("" & mytablex.Fields("producto"))
        mytabley.Fields("unidad1") = "UND"
        mytabley.Fields("factor1") = 1
        mytabley.Update
        'MsgBox "No existe Precio venta en dicho Local ", 48, "Aviso"
        'mytablex.Close
        mytabley.Close
        GoTo a134

    End If

    Data2.Recordset.AddNew
    graba_temporaldsi mytablex, mytabley, sw, canti
    Data2.Recordset.Update
    sw1 = 1
    busca_productosi = 1
    mytablex.Close
        
    'If sw1 = 1 And Len(ybuf) > 0 Then
    'If valida_flag("" & racu) = 2 Then    'compras
    '   found = crea_nuevos_proveedores("" & codigo, "" & xbuf, "" & ybuf)
    'End If
    'End If
    mytabley.Close

End Function

Sub proceso_v4()

    Dim buf      As String

    Dim mydby    As Database

    Dim mytabley As Snapshot

    Dim sdx      As Double

    Dim vr

    Dim mytablex As New ADODB.Recordset

    copiar_almacen0
    cn.Execute ("delete from almacen0 ")
    orionv4 = "\orion.v4\001d\01"
    sdx = 0
    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    buf = "select * from almacen where local='" & local1 & "' and bodega='" & bodega & "'"
    Set mytabley = mydby.CreateSnapshot(buf)
    mytablex.Open "SELECT * FROM almacen0", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew
        mytablex.Fields("producto") = Trim("" & mytabley.Fields("producto"))
        mytablex.Fields("local") = Trim(local1)
        mytablex.Fields("bodega") = Trim(bodega)
        mytablex.Fields("saldo") = Val("" & mytabley.Fields("SALDO"))
        mytablex.Update
        sdx = sdx + 1
        vr = DoEvents()
        mytabley.MoveNext
    Loop
    mytablex.Close
    mytabley.Close

    If mytablexsi.State = 1 Then mytablexsi.Close
    mytablexsi.Open "SELECT * FROM almacen0", cn, adOpenKeyset, adLockOptimistic
    Set DataGrid1.DataSource = mytablexsi
    sdx = mytablexsi.RecordCount
    Label41 = "" & sdx

    MsgBox "Proceso Terminado ", 48, "Aviso"

End Sub

Sub copiar_almacen0()

    On Error GoTo cmd90124_error

    cn.Execute ("select * into almacen0 from almacen where local='" & local1 & "'")
    Exit Sub
cmd90124_error:
    MsgBox "Almacen 0 ..ya existe puede continuar...", 48, "Aviso"
    Exit Sub

End Sub

Sub sql_controlpeso(buf As String)

    Dim sdx1 As Double

    Dim sdx2 As Double

    Dim sdx3 As Double

    Dim sdx4 As Double

    Dim sdx5 As Double

    Dim sdx6 As Double

    sdx1 = 0
    sdx2 = 0
    sdx3 = 0
    sdx4 = 0
    sdx5 = 0
    sdx6 = 0
    'If mytablepeso.State = 1 Then
    '   mytablepeso.Close
    'End If
    Set mytablepeso = Nothing
    mytablepeso.Open "SELECT * FROM controlpeso where producto='" & buf & "'", cn, adOpenKeyset, adLockOptimistic
    Set dbgrid13.DataSource = mytablepeso
    Do

        If mytablepeso.EOF Then Exit Do
        sdx1 = sdx1 + Val("" & mytablepeso.Fields("nrojabas"))
        sdx2 = sdx2 + Val("" & mytablepeso.Fields("cantprod"))
        sdx3 = sdx3 + Val("" & mytablepeso.Fields("pesobruto"))
        sdx4 = sdx4 + Val("" & mytablepeso.Fields("tara"))
        sdx5 = sdx5 + Val("" & mytablepeso.Fields("pesoneto"))
        sdx6 = sdx6 + Val("" & mytablepeso.Fields("total"))
        mytablepeso.MoveNext
    Loop
    nsdx1 = Format(sdx1, "0.00")
    nsdx2 = Format(sdx2, "0.00")
    nsdx3 = Format(sdx3, "0.00")
    nsdx4 = Format(sdx4, "0.00")
    nsdx5 = Format(sdx5, "0.00")
    nsdx6 = Format(sdx6, "0.00")

End Sub

Sub consulta_producto_inventario()

    Dim sdx      As Double

    Dim sdx1     As Double

    Dim mytabley As New ADODB.Recordset

    Set mytablexsi = Nothing
    mytablexsi.Open "SELECT * FROM producto where seinventaria='S' order by descripcio", cn, adOpenKeyset, adLockOptimistic
    sdx1 = mytablexsi.RecordCount

    If sdx1 > 0 Then
        If MsgBox("Desea Cargar Son " & sdx1 & " Registros", 1, "Aviso") <> 1 Then
            mytablexsi.Close
            Exit Sub

        End If

    End If

    Do

        If mytablexsi.EOF Then Exit Do

a1345:
        Set mytabley = Nothing
        mytabley.Open "SELECT * FROM precios where  producto='" & "" & mytablexsi.Fields("producto") & "' and local='" & local1 & "'", cn, adOpenKeyset, adLockOptimistic

        If mytabley.RecordCount = 0 Then  'si existe
            mytabley.AddNew
            mytabley.Fields("local") = Trim("" & local1)
            mytabley.Fields("producto") = Trim("" & mytablexsi.Fields("producto"))
            mytabley.Fields("unidad1") = "UND"
            mytabley.Fields("factor1") = 1
            mytabley.Update
            'MsgBox "No existe Precio venta en dicho Local ", 48, "Aviso"
            'mytablex.Close
            mytabley.Close
            GoTo a1345

        End If

        Data2.Recordset.AddNew
        graba_temporaldsi mytablexsi, mytabley, 0, 0
        Data2.Recordset.Update
        mytabley.Close
        mytablexsi.MoveNext
    Loop
    mytablexsi.Close
    sql_detalle
    sumar_detalle

End Sub

'25/06/2018 Testing Almacen General
Function ve_descarga2(buf As String, buf1 As String)

    Dim mytablex   As New ADODB.Recordset

    Dim mytablexyz As New ADODB.Recordset

    Dim acu        As String

    Dim acu1       As String

    acu = ""
    acu1 = ""

    mytablex.Open "SELECT * FROM tipo where  tipo='" & buf1 & "'", cn, adOpenKeyset, adLockOptimistic
    mytablexyz.Open "SELECT * FROM tipo where  tipo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        acu = "" & mytablexyz.Fields("tipodoc")
        acu1 = "" & mytablex.Fields("tipodoc")
    
        Select Case "" & acu

            Case "A", "B", "C", "D", "G", "E", "F", "J", "K", "L", "M", "P", "N", "O", "S", "T"

                If (acu1 = "A" Or acu1 = "B" Or acu1 = "C" Or acu1 = "D") And acu = "T" Then
                    ve_descarga2 = 1
                Else
                    ve_descarga2 = 0

                End If
                  
        End Select

    End If

    mytablex.Close

End Function

'25/06/2018 Testing Almacen General

'25/06/2018 Testing Almacen General

