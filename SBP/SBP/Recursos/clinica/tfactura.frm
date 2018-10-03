VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form tfactura 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla Facturas"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   14370
   LinkTopic       =   "Form1"
   ScaleHeight     =   9075
   ScaleWidth      =   14370
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
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
      Top             =   7680
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Consulta"
      Height          =   8895
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   13695
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
         Left            =   8280
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   600
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   600
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
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   600
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid dbGrid2 
         Height          =   7575
         Left            =   240
         TabIndex        =   43
         Top             =   1200
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   13361
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Frame1"
      Height          =   8895
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   13575
      Begin VB.TextBox transporte 
         Height          =   375
         Left            =   6360
         MaxLength       =   11
         TabIndex        =   14
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox bodega 
         Height          =   375
         Left            =   6360
         MaxLength       =   11
         TabIndex        =   15
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox turno 
         Height          =   375
         Left            =   3960
         MaxLength       =   1
         TabIndex        =   12
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox caja 
         Height          =   375
         Left            =   3960
         MaxLength       =   2
         TabIndex        =   5
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox cajero 
         Height          =   375
         Left            =   3960
         MaxLength       =   11
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox ruc 
         Height          =   375
         Left            =   1080
         MaxLength       =   11
         TabIndex        =   7
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "FormaPago"
         Height          =   6615
         Left            =   9960
         TabIndex        =   54
         Top             =   480
         Width           =   3495
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
            Left            =   2040
            MaskColor       =   &H00FFFFFF&
            Picture         =   "tfactura.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   5400
            UseMaskColor    =   -1  'True
            Width           =   1215
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
            Left            =   2040
            MaskColor       =   &H00FFFFFF&
            Picture         =   "tfactura.frx":07AE
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   4560
            UseMaskColor    =   -1  'True
            Width           =   1215
         End
         Begin VB.TextBox fpago 
            Height          =   375
            Left            =   120
            MaxLength       =   2
            TabIndex        =   60
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox pago1 
            Height          =   375
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   59
            Top             =   1800
            Width           =   855
         End
         Begin VB.TextBox fpago2 
            Height          =   375
            Left            =   120
            MaxLength       =   2
            TabIndex        =   58
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox pago2 
            Height          =   375
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   57
            Top             =   2160
            Width           =   855
         End
         Begin VB.TextBox fpago3 
            Height          =   375
            Left            =   120
            MaxLength       =   2
            TabIndex        =   56
            Top             =   2520
            Width           =   615
         End
         Begin VB.TextBox pago3 
            Height          =   375
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   55
            Top             =   2520
            Width           =   855
         End
         Begin VB.Label Label15 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fpago"
            Height          =   375
            Left            =   120
            TabIndex        =   76
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nombre"
            Height          =   375
            Left            =   720
            TabIndex        =   75
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label10 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Valor"
            Height          =   375
            Left            =   2400
            TabIndex        =   74
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label nfpago1 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   720
            TabIndex        =   73
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label nfpago2 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   720
            TabIndex        =   72
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label nfpago3 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   720
            TabIndex        =   71
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label Label20 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vuelto"
            Height          =   375
            Left            =   1680
            TabIndex        =   70
            Top             =   3240
            Width           =   1575
         End
         Begin VB.Label vuelto 
            BackColor       =   &H00FFFF00&
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
            Left            =   1680
            TabIndex        =   69
            Top             =   3600
            Width           =   1575
         End
         Begin VB.Label Label14 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "T/C"
            Height          =   375
            Left            =   120
            TabIndex        =   68
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label paridadf 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   720
            TabIndex        =   67
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label moneda3 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   2160
            TabIndex        =   66
            Top             =   2520
            Width           =   255
         End
         Begin VB.Label moneda2 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   2160
            TabIndex        =   65
            Top             =   2160
            Width           =   255
         End
         Begin VB.Label moneda1 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   2160
            TabIndex        =   64
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label Label17 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "M"
            Height          =   375
            Left            =   2160
            TabIndex        =   63
            Top             =   1440
            Width           =   255
         End
      End
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "tfactura.frx":0F5C
         Height          =   5295
         Left            =   120
         OleObjectBlob   =   "tfactura.frx":0F70
         TabIndex        =   17
         Top             =   2520
         Width           =   9735
      End
      Begin VB.TextBox sede 
         BackColor       =   &H00FFFF00&
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         MaxLength       =   2
         TabIndex        =   18
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox vendedor 
         Height          =   375
         Left            =   6360
         MaxLength       =   11
         TabIndex        =   13
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox fecha 
         Height          =   375
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox observa 
         BackColor       =   &H00FFFF00&
         Height          =   375
         Left            =   6360
         MaxLength       =   60
         TabIndex        =   16
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox paridad 
         Height          =   375
         Left            =   3960
         MaxLength       =   11
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox moneda 
         Height          =   375
         Left            =   3960
         MaxLength       =   1
         TabIndex        =   9
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox codigo 
         Height          =   375
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   6
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox tipocliente 
         Height          =   375
         Left            =   1080
         MaxLength       =   11
         TabIndex        =   4
         Text            =   "P"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox numero 
         Height          =   375
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox serie 
         Height          =   375
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   2
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox tipo 
         Height          =   375
         Left            =   600
         MaxLength       =   2
         TabIndex        =   1
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja"
         Height          =   375
         Left            =   3000
         TabIndex        =   83
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Transportista"
         Height          =   375
         Left            =   5400
         TabIndex        =   82
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Almacen"
         Height          =   375
         Left            =   5400
         TabIndex        =   81
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label ncodigo 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   80
         Top             =   7920
         Width           =   3015
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Turno"
         Height          =   375
         Left            =   3000
         TabIndex        =   79
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cajero"
         Height          =   375
         Left            =   3000
         TabIndex        =   78
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ruc"
         Height          =   375
         Left            =   120
         TabIndex        =   77
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label xsaldo 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1080
         TabIndex        =   53
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saldo"
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label cantidad 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3120
         TabIndex        =   51
         Top             =   7920
         Width           =   975
      End
      Begin VB.Label estadocredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   360
         TabIndex        =   50
         Top             =   8280
         Width           =   255
      End
      Begin VB.Label estado 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   8280
         Width           =   255
      End
      Begin VB.Label abono 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3120
         TabIndex        =   48
         Top             =   8280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label saldo 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   47
         Top             =   8280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label nvendedor 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1800
         TabIndex        =   46
         Top             =   8280
         Width           =   1335
      End
      Begin VB.Label ntipo 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   600
         TabIndex        =   45
         Top             =   8280
         Width           =   1215
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         Height          =   375
         Left            =   600
         TabIndex        =   44
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vendedor"
         Height          =   375
         Left            =   5400
         TabIndex        =   38
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha"
         Height          =   375
         Left            =   5400
         TabIndex        =   37
         Top             =   600
         Width           =   975
      End
      Begin VB.Label total 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8040
         TabIndex        =   36
         Top             =   7920
         Width           =   1815
      End
      Begin VB.Label impuesto 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6960
         TabIndex        =   35
         Top             =   7920
         Width           =   1095
      End
      Begin VB.Label subtotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6000
         TabIndex        =   34
         Top             =   7920
         Width           =   975
      End
      Begin VB.Label descuento 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5040
         TabIndex        =   33
         Top             =   7920
         Width           =   975
      End
      Begin VB.Label neto 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   32
         Top             =   7920
         Width           =   975
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observacion"
         Height          =   375
         Left            =   5400
         TabIndex        =   31
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "T/C"
         Height          =   375
         Left            =   3000
         TabIndex        =   30
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Moneda"
         Height          =   375
         Left            =   3000
         TabIndex        =   29
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cliente(P/A)"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie  Numero"
         Height          =   375
         Left            =   1080
         TabIndex        =   27
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sede"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFF00&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   14310
      TabIndex        =   19
      Top             =   0
      Width           =   14370
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
         Left            =   9240
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tfactura.frx":3233
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox xsede 
         Height          =   375
         Left            =   5040
         MaxLength       =   6
         TabIndex        =   87
         Top             =   0
         Width           =   1215
      End
      Begin VB.TextBox xcliente 
         Height          =   375
         Left            =   5040
         MaxLength       =   11
         TabIndex        =   86
         Text            =   "*"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox fechai 
         Height          =   375
         Left            =   7440
         MaxLength       =   10
         TabIndex        =   85
         Top             =   0
         Width           =   1695
      End
      Begin VB.TextBox fechaf 
         Height          =   375
         Left            =   7440
         MaxLength       =   10
         TabIndex        =   84
         Top             =   360
         Width           =   1695
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
         Left            =   1440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tfactura.frx":39E1
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Ayuda"
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
         Picture         =   "tfactura.frx":4BF3
         Style           =   1  'Graphical
         TabIndex        =   23
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
         Height          =   615
         Left            =   2880
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tfactura.frx":5E05
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Height          =   615
         Left            =   2160
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tfactura.frx":7017
         Style           =   1  'Graphical
         TabIndex        =   21
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
         Height          =   615
         Left            =   0
         Picture         =   "tfactura.frx":8229
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sede"
         Height          =   375
         Left            =   3720
         TabIndex        =   91
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cliente"
         Height          =   375
         Left            =   3720
         TabIndex        =   90
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   375
         Left            =   6240
         TabIndex        =   89
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Final"
         Height          =   375
         Left            =   6240
         TabIndex        =   88
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid dbgrid1 
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   14208
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
   Begin VB.Menu ahyy1 
      Caption         =   "&Add"
   End
   Begin VB.Menu dmi22 
      Caption         =   "&Modifica"
   End
   Begin VB.Menu dfj8221 
      Caption         =   "&Borra"
   End
   Begin VB.Menu dk281 
      Caption         =   "&Reporte"
   End
   Begin VB.Menu fdo33 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tfactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim td1 As New ADODB.Recordset
Dim rsfac As New ADODB.Recordset
Public rt As New ADODB.Recordset

Private Sub sql()
On Error GoTo cmd5_err
Dim cad As String
If Len(xsede) = 0 Then
   MsgBox "Sede no Existe ", 48, "Aviso"
   xsede.SetFocus
   Exit Sub
End If
If Not IsDate(fechai) Then
   MsgBox "Fecha Erronea", 48, "Aviso"
   fechai.SetFocus
   Exit Sub
End If
If Not IsDate(fechaf) Then
   MsgBox "Fecha Erronea", 48, "Aviso"
   fechaf.SetFocus
   Exit Sub
End If
cad = "SELECT sede,Tipo,Serie,Numero,Cliente,Nombre,Tipocliente as 'T',Moneda as 'M',Paridad,Formapago as 'Fp',Neto,Descuento,Totalfac,Impuesto,Subtotal,Abono,Saldo,Estado,EstadoCredito,vendedor,fecha,observa,fpago1,fpago2,fpago3,moneda1,moneda2,moneda3,paga1,paga2,paga3,paridadf,ruc,cajero,caja,turno,transporte,bodega FROM factura "
cad = cad & "where sede='" & xsede & "'"
cad = cad & " and fecha>='" & Format(fechai, "YYYYMMDD") & "'"
cad = cad & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "'"
cad = cad & " order by fecha"
If rsfac.State = 1 Then rsfac.Close
rsfac.Open cad, cn, adOpenStatic, adLockOptimistic
Set dbgrid1.DataSource = rsfac
dbgrid1.Columns(0).Width = 500
dbgrid1.Columns(1).Width = 500
dbgrid1.Columns(2).Width = 500
dbgrid1.Columns(3).Width = 800
dbgrid1.Columns(4).Width = 1000
dbgrid1.Columns(5).Width = 3000
dbgrid1.Columns(6).Width = 300
dbgrid1.Columns(7).Width = 800
dbgrid1.Columns(8).Width = 300
dbgrid1.Columns(9).Width = 800
dbgrid1.Columns(10).Width = 800
dbgrid1.Columns(11).Width = 800
dbgrid1.Columns(12).Width = 800
dbgrid1.Columns(13).Width = 800
dbgrid1.Columns(14).Width = 800
dbgrid1.Columns(15).Width = 800
dbgrid1.Columns(16).Width = 800
dbgrid1.Columns(17).Width = 800
dbgrid1.Columns(18).Width = 800

Exit Sub
cmd5_err:
MsgBox "Aviso en sql " + Error, 48, "Aviso"
Exit Sub
End Sub

Private Sub ahyy1_Click()
Dim found As Integer
If Frame1.Visible = True Then Exit Sub
If existe_sede("" & xsede) = 0 Then
   MsgBox "Sede no existe ", 48, "Aviso"
   xsede.SetFocus
   Exit Sub
End If

Frame1.Visible = True
Frame1.Caption = "NUEVO"
tipo = ""
serie = ""
numero = ""
sede = xsede
carga_detalle
carga_detalle
carga_detalle_borra

tipo.Enabled = True
serie.Enabled = True
numero.Enabled = True
tipocliente.Enabled = True
codigo.Enabled = True
'caja.Enabled = True
paridadf = "1"
paridadf = "" & busca_paridadf()
If Val(paridadf) <= 0 Then
   paridadf = "1"
End If
inicializa
suma_fpago
sede = gsede1
tipo.SetFocus
End Sub
Sub inicializa()
cajero = ""
caja = ""
turno = ""
transporte = ""
bodega = ""
ruc = ""
ncodigo = ""
saldo = ""
paridadf = "1"
vuelto = ""
Label20 = ""
tipocliente = ""
codigo = ""
fecha = Format(Now, "dd/mm/yyyy")
moneda = "S"
paridad = ""
fpago2 = ""
fpago3 = ""
fpago = ""
vendedor = ""
observa = ""
moneda1 = ""
moneda2 = ""
moneda3 = ""
pago1 = ""
pago3 = ""
pago2 = ""
vuelto = ""
End Sub

Private Sub bodega_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
observa.SetFocus

End Sub

Private Sub bodega_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_bodega
End If

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   If opcion1 = 1 Then
      Frame2.Visible = False
      tipo.SetFocus
      Exit Sub
   End If
   If opcion1 = 2 Then
      Frame2.Visible = False
      codigo.SetFocus
      Exit Sub
   End If
   If opcion1 = 3 Then
      Frame2.Visible = False
      fpago.SetFocus
      Exit Sub
   End If
   If opcion1 = 4 Then
      Frame2.Visible = False
      vendedor.SetFocus
      Exit Sub
   End If
   If opcion1 = 30 Then
      Frame2.Visible = False
      cajero.SetFocus
      Exit Sub
   End If
   If opcion1 = 31 Then
      Frame2.Visible = False
      caja.SetFocus
      Exit Sub
   End If
   If opcion1 = 32 Then
      Frame2.Visible = False
      transporte.SetFocus
      Exit Sub
   End If
   
   If opcion1 = 5 Then
      Frame2.Visible = False
      DBGrid3.SetFocus
      Exit Sub
   End If
   If opcion1 = 50 Then
      Frame2.Visible = False
      DBGrid3.SetFocus
      Exit Sub
   End If
   If opcion1 = 51 Then
      Frame2.Visible = False
      DBGrid3.SetFocus
      Exit Sub
   End If
      
      If opcion1 = 7 Then
      Frame2.Visible = False
      fpago2.SetFocus
      Exit Sub
   End If
      If opcion1 = 8 Then
      Frame2.Visible = False
      fpago3.SetFocus
      Exit Sub
   End If



End If

End Sub

Private Sub caja_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If existe_caja("" & caja) = 0 Then
   MsgBox "Caja no existe", 48, "Aviso"
   caja.SetFocus
   Exit Sub
End If
turno.SetFocus

End Sub

Private Sub caja_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_caja
End If

End Sub

Private Sub cajero_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
caja.SetFocus
End Sub

Private Sub cajero_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_cajero
End If

End Sub

Private Sub cmdAddEntry_Click()
ahyy1_Click
End Sub

Private Sub cmdDelete_Click()
dfj8221_Click
End Sub

Private Sub cmdExit_Click()
fdo33_Click
End Sub

Private Sub cmdHelp_Click()
dmi22_Click
End Sub

Private Sub cmdPrint_Click()
dk281_Click
End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(codigo) = 0 Then
   codigo.SetFocus
   Exit Sub
End If
found = existe_codigo()
If found = 0 Then
   codigo.SetFocus
   Exit Sub
End If
ncodigo = codigo_nombre("" & codigo)
consulta_saldo
ruc.SetFocus
End Sub

Private Sub codigo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_codigo
End If

End Sub

Private Sub Command1_Click()
ejecuta 1
End Sub

Private Sub Command2_Click()
sql
End Sub

Private Sub Command3_Click()
Dim found As Integer
Dim rsexiste As New ADODB.Recordset
Dim cad As String
On Error GoTo cmd2_err

If Len(sede) = 0 Then
   sede.SetFocus
   Exit Sub
End If
If Len(tipo) = 0 Then
   tipo.SetFocus
   Exit Sub
End If

If existe_tipo() = 0 Then
   MsgBox "Tipo Documento no existe", 48, "Aviso"
   tipo.SetFocus
   Exit Sub
End If
If tipo = "2" Or tipo = "4" Then
   If Len(ruc) = 0 Then
      MsgBox "Ruc no valido ", 48, "Aviso"
      ruc.SetFocus
      Exit Sub
   End If
End If

If Len(serie) = 0 Then
   serie.SetFocus
   Exit Sub
End If
If Len(numero) = 0 Then
   numero.SetFocus
   Exit Sub
End If
If tipocliente <> "A" And tipocliente <> "P" Then
   tipocliente.SetFocus
   Exit Sub
End If
If Len(codigo) = 0 Then
   codigo.SetFocus
   Exit Sub
End If
If existe_codigo() = 0 Then
   MsgBox "Codigo no existe", 48, "Aviso"
   codigo.SetFocus
   Exit Sub
End If

If Not IsDate(fecha) Then
   fecha.SetFocus
   Exit Sub
End If
If moneda <> "S" And moneda <> "D" Then
   moneda.SetFocus
   Exit Sub
End If
If Len(vendedor) = 0 Then
   vendedor.SetFocus
   Exit Sub
End If

If existe_cajero("" & cajero) = 0 Then
   MsgBox "Usuario no existe", 48, "Aviso"
   cajero.SetFocus
   Exit Sub
End If
If existe_caja("" & caja) = 0 Then
   MsgBox "Caja no existe", 48, "Aviso"
   caja.SetFocus
   Exit Sub
End If
If existe_turno("" & turno) = 0 Then
   MsgBox "Caja no existe", 48, "Aviso"
   turno.SetFocus
   Exit Sub
End If
If Len(transporte) > 0 Then
If existe_transporte("" & transporte) = 0 Then
   MsgBox "Transporte no existe", 48, "Aviso"
   transporte.SetFocus
   Exit Sub
End If
End If
If existe_bodega("" & bodega) = 0 Then
   MsgBox "Almacen no existe", 48, "Aviso"
   bodega.SetFocus
   Exit Sub
End If


If Len(fpago) = 0 Then
   fpago.SetFocus
   Exit Sub
End If

suma_fpago
If Val(vuelto) < Val(total) And Val(vuelto) > 0 Then
   MsgBox "Falta Poner Forma Pago ", 48, "Aviso"
   fpago.SetFocus
   Exit Sub
End If


If Frame1.Caption = "NUEVO" Then
   If existe_numero() = 1 Then
      MsgBox "Ya existe dicho numero ", 48, "Aviso"
      numero.SetFocus
      Exit Sub
   End If
   cad = "INSERT INTO factura VALUES('" & Trim(tipo) & "','" & Trim(serie) & "','" & Trim(numero) & "','" & Trim(codigo) & "','" & Trim(tipocliente) & "','" & Trim(moneda) & "'," & Val(paridad) & ",'" & Trim(fpago) & "'," & Val(neto) & "," & Val(descuento) & "," & Val(total) & "," & Val(impuesto) & "," & Val(subtotal) & "," & Val(abono) & "," & Val(saldo) & ",'" & Trim(estado) & "','" & Trim(estadocredito) & "','" & Trim(sede) & "','" & Trim(vendedor) & "','" & Trim(fecha) & "','" & Trim(observa) & "','" & Trim(fpago) & "','" & Trim(fpago2) & "','" & Trim(fpago3) & "'," & Val(pago1) & "," & Val(pago2) & ",'" & Val(pago3) & "','" & Trim(moneda1) & "','" & Trim(moneda2) & "','" & Trim(moneda3) & "'," & Val(paridadf) & ",'" & Trim(ruc) & "','" & Trim(cajero) & "','" & Trim(caja) & "','" & Trim(turno) & "','" & Trim(transporte) & "','" & Trim(bodega) & "','" & Trim(ncodigo) & "')"
   cn.Execute (cad)
   grabar_detalle
   sql
   dbgrid1.SetFocus
   fdo33_Click
End If
If Frame1.Caption = "MODIFICA" Then
   cad = "UPDATE factura SET cliente = '" & Trim(codigo) & "', tipocliente= '" & Trim(tipocliente) & "', moneda= '" & Trim(moneda) & "', paridad= " & Val(paridad) & ", formapago= '" & Trim(fpago) & "', neto= " & Val(neto) & ", descuento= " & Val(descuento) & ", totalfac= " & Val(total) & ", impuesto= " & Val(impuesto) & ", subtotal= " & Val(subtotal) & ", abono= " & Val(abono) & ", saldo= " & Val(saldo) & ", estado= '" & Trim(estado) & "', estadocredito= '" & Trim(estadocredito) & "', vendedor= '" & Trim(vendedor) & "', fecha= '" & Trim(fecha) & "', observa= '" & Trim(observa) & "', fpago1= '" & Trim(fpago) & "', fpago2= '" & Trim(fpago2) & "'  , fpago3= '" & Trim(fpago3) & "', paga1= " & Val(pago1) & ", paga2= " & Val(pago2) & " , paga3= " & Val(pago3) & ", moneda1= '" & Trim(moneda1) & "', moneda2= '" & Trim(moneda2) & "', moneda3= '" & Trim(moneda3) & "', paridadf= " & Val(paridadf) & ",ruc='"
   cad = cad & Trim(ruc) & "',cajero='" & Trim(cajero) & "',caja='" & Trim(caja) & "' , turno='" & Trim(turno) & "',transporte='" & Trim(transporte) & "',bodega='" & Trim(bodega) & "',nombre='" & Trim(ncodigo) & "' WHERE sede = '" & Trim(sede) & "' and tipo='" & tipo & "' and serie='" & serie & "' and numero='" & numero & "'"
   cn.Execute (cad)
   cn.Execute ("DELETE   FROM detalle WHERE sede ='" & Trim(sede) & "' and tipo='" & Trim(tipo) & "' and serie='" & Trim(serie) & "' and numero='" & Trim(numero) & "'")
   grabar_detalle
   sql
   dbgrid1.SetFocus
   fdo33_Click
End If


Exit Sub
cmd2_err:
MsgBox "Aviso en command3 " + Error$, 48, "Aviso"
Exit Sub

End Sub

Private Sub Command4_Click()
fdo33_Click
End Sub

Private Sub dbGrid2_KeyPress(KeyAscii As Integer)
Dim buf As String
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
         ejecuta 0
End If

End Sub

Private Sub dbGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
   If opcion1 = 1 Then
      tipo = Trim(dbGrid2.Columns(1))
      ntipo = Trim(dbGrid2.Columns(0))
      Frame2.Visible = False
      Frame2.Enabled = False
      serie.SetFocus
      Exit Sub
   End If
   If opcion1 = 2 Then
      codigo = Trim(dbGrid2.Columns(1))
      ncodigo = Trim(dbGrid2.Columns(0))
      ruc = Trim(dbGrid2.Columns(2))
      Frame2.Visible = False
      Frame2.Enabled = False
      codigo.SetFocus
      Exit Sub
   End If
   If opcion1 = 3 Then
      fpago = Trim(dbGrid2.Columns(1))
      nfpago1 = Trim(dbGrid2.Columns(0))
      moneda1 = Trim(dbGrid2.Columns(2))
      Frame2.Visible = False
      Frame2.Enabled = False
      pago1.SetFocus
      Exit Sub
   End If
   If opcion1 = 7 Then
      fpago2 = Trim(dbGrid2.Columns(1))
      nfpago2 = Trim(dbGrid2.Columns(0))
      moneda2 = Trim(dbGrid2.Columns(2))
      Frame2.Visible = False
      Frame2.Enabled = False
      pago2.SetFocus
      Exit Sub
   End If
   If opcion1 = 8 Then
      fpago3 = Trim(dbGrid2.Columns(1))
      nfpago3 = Trim(dbGrid2.Columns(0))
      moneda3 = Trim(dbGrid2.Columns(2))
      Frame2.Visible = False
      Frame2.Enabled = False
      pago3.SetFocus
      Exit Sub
   End If

   
   If opcion1 = 4 Then
      vendedor = Trim(dbGrid2.Columns(1))
      nvendedor = Trim(dbGrid2.Columns(0))
      Frame2.Visible = False
      Frame2.Enabled = False
      vendedor.SetFocus
      Exit Sub
   End If
   If opcion1 = 30 Then
      cajero = Trim(dbGrid2.Columns(1))
      Frame2.Visible = False
      Frame2.Enabled = False
      cajero.SetFocus
      Exit Sub
   End If
   If opcion1 = 31 Then
      caja = Trim(dbGrid2.Columns(1))
      Frame2.Visible = False
      Frame2.Enabled = False
      caja.SetFocus
      Exit Sub
   End If
   If opcion1 = 32 Then
      transporte = Trim(dbGrid2.Columns(1))
      Frame2.Visible = False
      Frame2.Enabled = False
      transporte.SetFocus
      Exit Sub
   End If
      If opcion1 = 33 Then
      bodega = Trim(dbGrid2.Columns(1))
      Frame2.Visible = False
      Frame2.Enabled = False
      bodega.SetFocus
      Exit Sub
   End If

   
   If opcion1 = 5 Then  'busa producto
      dbGrid2.Refresh
      DBGrid3.Columns(0) = Trim(dbGrid2.Columns(1))
      DBGrid3.Columns(1) = Trim(dbGrid2.Columns(0))
      DBGrid3.Columns(2) = "UND"
      DBGrid3.Columns(3) = 1
      DBGrid3.Columns(4) = 1
      DBGrid3.Columns(5) = Val(Trim(dbGrid2.Columns(2)))
      DBGrid3.Columns(6) = 0
      DBGrid3.Columns(7) = Val("" & DBGrid3.Columns(4)) * Val("" & DBGrid3.Columns(5))
      DBGrid3.Columns(12) = Val(Trim(dbGrid2.Columns(3)))
      calcula_datos
      Frame2.Visible = False
      Frame2.Enabled = False
      sumar_detalle
      DBGrid3.Col = 0
      DBGrid3.Row = DBGrid3.VisibleRows - 1
      DBGrid3.SetFocus
      
      Exit Sub
   End If
   If opcion1 = 50 Then  'Consulta a pagar
      dbGrid2.Refresh
      'consulta , Partsaldo, Segusaldo, sede, cliente
      DBGrid3.Columns(14) = Trim(dbGrid2.Columns(0))  'consulta
      DBGrid3.Columns(15) = Trim(dbGrid2.Columns(3))  'sede
      DBGrid3.Columns(16) = "C"  'consulta
      DBGrid3.Columns(0) = "PC"
      DBGrid3.Columns(1) = "CONSULTA "
      DBGrid3.Columns(2) = "UND"
      DBGrid3.Columns(3) = 1
      DBGrid3.Columns(4) = 1
      If tipocliente = "P" Then
      DBGrid3.Columns(5) = Val(Trim(dbGrid2.Columns(1)))
      End If
      If tipocliente = "A" Then
      DBGrid3.Columns(5) = Val(Trim(dbGrid2.Columns(2)))
      End If
      DBGrid3.Columns(6) = 0
      DBGrid3.Columns(7) = Val("" & DBGrid3.Columns(4)) * Val("" & DBGrid3.Columns(5))
      DBGrid3.Columns(12) = Val(Trim(dbGrid2.Columns(3)))
      calcula_datos
      Frame2.Visible = False
      Frame2.Enabled = False
      sumar_detalle
      DBGrid3.Col = 0
      DBGrid3.Row = DBGrid3.VisibleRows - 1
      DBGrid3.SetFocus
      
      Exit Sub
   End If
   If opcion1 = 51 Then  'tratamiento a pagar
      dbGrid2.Refresh
      'Tratamiento,PagaParticular,PagaEmpresa,Sede,Cliente ,Consulta
      DBGrid3.Columns(13) = Trim(dbGrid2.Columns(0))  'consulta
      DBGrid3.Columns(14) = Trim(dbGrid2.Columns(5))  'consulta
      DBGrid3.Columns(15) = Trim(dbGrid2.Columns(3))  'sede
      DBGrid3.Columns(16) = "T"  'tratamiento
      DBGrid3.Columns(0) = "PT"
      DBGrid3.Columns(1) = "TRATAMIENTO"
      DBGrid3.Columns(2) = "UND"
      DBGrid3.Columns(3) = 1
      DBGrid3.Columns(4) = 1
      If tipocliente = "P" Then
      DBGrid3.Columns(5) = Val(Trim(dbGrid2.Columns(1)))
      End If
      If tipocliente = "A" Then
      DBGrid3.Columns(5) = Val(Trim(dbGrid2.Columns(2)))
      End If
      DBGrid3.Columns(6) = 0
      DBGrid3.Columns(7) = Val("" & DBGrid3.Columns(4)) * Val("" & DBGrid3.Columns(5))
      DBGrid3.Columns(12) = Val(Trim(dbGrid2.Columns(3)))
      calcula_datos
      Frame2.Visible = False
      Frame2.Enabled = False
      sumar_detalle
      DBGrid3.Col = 0
      DBGrid3.Row = DBGrid3.VisibleRows - 1
      DBGrid3.SetFocus
      
      Exit Sub
   End If

End If

End Sub

Private Sub DBGrid3_AfterColUpdate(ByVal ColIndex As Integer)
Select Case ColIndex
       Case 0
            DBGrid3.Col = 0
            DBGrid3.Row = DBGrid3.VisibleRows - 1
            DBGrid3.SetFocus
       Case 4
            DBGrid3.Col = 0
            DBGrid3.Row = DBGrid3.VisibleRows - 1
            DBGrid3.SetFocus
End Select
End Sub

Private Sub DBGrid3_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
Select Case ColIndex
     Case 1, 2, 3, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17
     Cancel = True
     Case 0
     If Len("" & DBGrid3.Columns(0)) > 0 Then  'si ya existe no cambiar
        Cancel = True
        Exit Sub
     End If
     Case 4
     If Len("" & DBGrid3.Columns(0)) = 0 Then  '
        Cancel = True
        Exit Sub
     End If
     Case 5
     If Len("" & DBGrid3.Columns(0)) = 0 Then  '
        Cancel = True
        Exit Sub
     End If
     Case 6
     If Len("" & DBGrid3.Columns(0)) = 0 Then  '
        Cancel = True
        Exit Sub
     End If
     Case 7
     If Len("" & DBGrid3.Columns(0)) = 0 Then  '
        Cancel = True
        Exit Sub
     End If
     

     
     Exit Sub
End Select

End Sub

Private Sub DBGrid3_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Select Case ColIndex
Case 0
     If Len(DBGrid3.Columns(0)) = 0 Then
        Cancel = True
        Exit Sub
     End If
     If Len(DBGrid3.Columns(0)) > 15 Then
        Cancel = True
        Exit Sub
     End If
Case 4
     If Len(DBGrid3.Columns(0)) = 0 Then
        Cancel = True
        Exit Sub
     End If
     If Not IsNumeric(DBGrid3.Columns(4)) Then
        Cancel = True
        Exit Sub
     End If
     calcula_datos
            DBGrid3.Col = 0
            DBGrid3.Row = DBGrid3.VisibleRows - 1
            DBGrid3.SetFocus
Case 5
     If Len(DBGrid3.Columns(0)) = 0 Then
        Cancel = True
        Exit Sub
     End If
     If Not IsNumeric(DBGrid3.Columns(5)) Then
        Cancel = True
        Exit Sub
     End If
     calcula_datos
            DBGrid3.Col = 0
            DBGrid3.Row = DBGrid3.VisibleRows - 1
            DBGrid3.SetFocus
Case 6
     If Len(DBGrid3.Columns(0)) = 0 Then
        Cancel = True
        Exit Sub
     End If
     If Not IsNumeric(DBGrid3.Columns(6)) Then
        Cancel = True
        Exit Sub
     End If
     calcula_datos
            DBGrid3.Col = 0
            DBGrid3.Row = DBGrid3.VisibleRows - 1
            DBGrid3.SetFocus

Case 7
     If Len(DBGrid3.Columns(0)) = 0 Then
        Cancel = True
        Exit Sub
     End If
     If Not IsNumeric(DBGrid3.Columns(7)) Then
        Cancel = True
        Exit Sub
     End If
     If Val(DBGrid3.Columns(4)) <= 0 Then
        Cancel = True
        Exit Sub
     End If
     If Val(DBGrid3.Columns(7)) <= 0 Then
        Cancel = True
        Exit Sub
     End If
     DBGrid3.Columns(5) = Val("" & DBGrid3.Columns(7)) / Val("" & DBGrid3.Columns(4))
     calcula_datos
            DBGrid3.Col = 0
            DBGrid3.Row = DBGrid3.VisibleRows - 1
            DBGrid3.SetFocus
End Select

End Sub
Sub consulta_consulta()
Dim cad As String
   Dim rconsulta As New ADODB.Recordset
   cad = "SELECT * FROM consulta where cliente='" & Trim(codigo) & "'"
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True Or rconsulta.BOF = True Then
      Exit Sub
   End If

Frame2.Visible = True
Frame2.Enabled = True
buffer = ""
Combo1.Clear
Combo1.AddItem "Consulta"
Combo1.ListIndex = 0
opcion1 = 50
buffer.SetFocus
Command1_Click
End Sub


Private Sub dbgrid3_KeyUp(KeyCode As Integer, Shift As Integer)
Dim found As Integer

If KeyCode = &H70 Then  'f1  carga los demas precios
   consulta_producto
End If
If KeyCode = &H71 Then  'f2  paga consulta
   consulta_consulta
End If
If KeyCode = &H72 Then  'f3  paga tratamiento
   consulta_tratamiento
End If


If KeyCode = &H2E Then  'borrar linea
If DBGrid3.Row = -1 Then
        MsgBox "No hay ningún registro para eliminar", vbInformation
        Exit Sub
End If
If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
   found = sumar_detalle()
   DBGrid3.Col = 0
   DBGrid3.Row = DBGrid3.VisibleRows - 1
   DBGrid3.Refresh
   DBGrid3.SetFocus
   Exit Sub
End If

If MsgBox("Se va a eliminar el registro : está seguro ", _
   vbExclamation + vbYesNo, "Eliminar") = vbYes Then
   Data1.Recordset.Delete
   If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
      found = sumar_detalle()
      DBGrid3.Col = 0
      DBGrid3.Row = DBGrid3.VisibleRows - 1
      DBGrid3.Refresh
      DBGrid3.SetFocus
      Exit Sub
   End If
   found = sumar_detalle()
   DBGrid3.Col = 0
   DBGrid3.Row = DBGrid3.VisibleRows - 1
   DBGrid3.SetFocus
   Exit Sub
End If
End If
End Sub

Private Sub dfj8221_Click()
Dim buf As String
On Error GoTo cmd4_err
If Frame1.Visible = True Then Exit Sub
buf = Trim(dbgrid1.Columns(1))
If MsgBox("Desea Borrar " + dbgrid1.Columns(1), 1, "Aviso") = 1 Then
   cn.Execute ("DELETE   FROM detalle WHERE sede ='" & Trim(dbgrid1.Columns(0)) & "' and tipo='" & Trim(dbgrid1.Columns(1)) & "' and serie='" & Trim(dbgrid1.Columns(2)) & "' and numero='" & Trim(dbgrid1.Columns(3)) & "'")
   cn.Execute ("DELETE   FROM factura WHERE sede ='" & Trim(dbgrid1.Columns(0)) & "' and tipo='" & Trim(dbgrid1.Columns(1)) & "' and serie='" & Trim(dbgrid1.Columns(2)) & "' and numero='" & Trim(dbgrid1.Columns(3)) & "'")
   rsfac.Requery
   sql
   dbgrid1.SetFocus
End If
dbgrid1.SetFocus
Exit Sub
cmd4_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
dbgrid1.SetFocus
Exit Sub

End Sub

Private Sub dk281_Click()
If Frame1.Visible = True Then Exit Sub
If rt.State = 1 Then rt.Close
rt.Open "SELECT * FROM sede ", cn, adOpenKeyset, adLockOptimistic
Set trepcli1.DataSource = rt
trepcli1.Show 1

End Sub

Private Sub dmi22_Click()
Dim found As Integer
On Error GoTo cmd3_err
If Frame1.Visible = True Then Exit Sub

sede = Trim(dbgrid1.Columns(0))
tipo = Trim(dbgrid1.Columns(1))
serie = Trim(dbgrid1.Columns(2))
numero = Trim(dbgrid1.Columns(3))
codigo = Trim(dbgrid1.Columns(4))
ncodigo = Trim(dbgrid1.Columns(5))
tipocliente = Trim(dbgrid1.Columns(6))
moneda = Trim(dbgrid1.Columns(7))
paridad = Trim(dbgrid1.Columns(8))
fpago = Trim(dbgrid1.Columns(9))
neto = Trim(dbgrid1.Columns(10))
descuento = Trim(dbgrid1.Columns(11))
total = Trim(dbgrid1.Columns(12))
impuesto = Trim(dbgrid1.Columns(13))
subtotal = Trim(dbgrid1.Columns(14))
abono = Trim(dbgrid1.Columns(15))
saldo = Trim(dbgrid1.Columns(16))
estado = Trim(dbgrid1.Columns(17))
estadocredito = Trim(dbgrid1.Columns(18))
vendedor = Trim(dbgrid1.Columns(19))
fecha = Trim(dbgrid1.Columns(20))
observa = Trim(dbgrid1.Columns(21))
fpago = Trim(dbgrid1.Columns(22))
fpago2 = Trim(dbgrid1.Columns(23))
fpago3 = Trim(dbgrid1.Columns(24))

nfpago1 = fpago_nombre(fpago)
nfpago2 = fpago_nombre(fpago2)
nfpago3 = fpago_nombre(fpago3)

moneda1 = Trim(dbgrid1.Columns(25))
moneda2 = Trim(dbgrid1.Columns(26))
moneda3 = Trim(dbgrid1.Columns(27))

pago1 = Trim(dbgrid1.Columns(28))
pago2 = Trim(dbgrid1.Columns(29))
pago3 = Trim(dbgrid1.Columns(30))
paridadf = Trim(dbgrid1.Columns(31))
ruc = Trim(dbgrid1.Columns(32))
cajero = Trim(dbgrid1.Columns(33))
caja = Trim(dbgrid1.Columns(34))
turno = Trim(dbgrid1.Columns(35))
transporte = Trim(dbgrid1.Columns(36))
bodega = Trim(dbgrid1.Columns(37))
'caja.Enabled = False
suma_fpago
found = existe_codigo()
If found = 0 Then
   codigo.SetFocus
   Exit Sub
End If
ncodigo = codigo_nombre("" & codigo)
consulta_saldo

Frame1.Visible = True
Frame1.Caption = "MODIFICA"
tipo.Enabled = False
serie.Enabled = False
numero.Enabled = False
tipocliente.Enabled = False
codigo.Enabled = False

carga_detalle
carga_detalle_datos
found = sumar_detalle()
ruc.SetFocus
Exit Sub
cmd3_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub
End Sub

Private Sub fdo33_Click()
If Frame2.Visible = True Then
   buffer_KeyPress 27
   Exit Sub
End If
If Frame1.Visible = True Then
   If Frame1.Caption = "NUEVO" Then
      Frame1.Visible = False
      dbgrid1.SetFocus
   End If
   If Frame1.Caption = "MODIFICA" Then
      Frame1.Visible = False
      dbgrid1.SetFocus
   End If
   Exit Sub
End If
tfactura.Hide
Unload tfactura
End Sub

Private Sub fecha_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
vendedor.SetFocus

End Sub

Private Sub Form_Load()
Dim found As Integer
'dgusuario1 = dgusuario
'MsgBox dgusuario
fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
fechaf = Format(Now, "dd/mm/yyyy")
xsede = gsede1
found = copiar_tmp("" & dgusuario)
If found = 0 Then
   MsgBox "Usuario en uso ", 48, "Aviso"
   Exit Sub
End If
sql

End Sub
Sub suma_fpago()
Dim sdx As Double
Dim sdx1 As Double
Dim xpago1 As Double
Dim xpago2 As Double
Dim xpago3 As Double
sdx1 = 1
paridadf = "" & busca_paridadf()
If Val(paridadf) <= 0 Then
   paridadf = "1"
End If

If Len(fpago) = 0 Then
   fpago.SetFocus
   Exit Sub
End If
xpago1 = Val(pago1)
xpago2 = Val(pago2)
xpago3 = Val(pago3)
If moneda = "S" Then
If moneda1 = "D" Then
   xpago1 = xpago1 * Val(paridadf)
End If
If moneda2 = "D" Then
   xpago2 = xpago2 * Val(paridadf)
End If
If moneda3 = "D" Then
   xpago3 = xpago3 * Val(paridadf)
End If
End If

If moneda = "D" Then
If moneda1 = "S" Then
   xpago1 = xpago1 / Val(paridadf)
End If
If moneda2 = "S" Then
   xpago2 = xpago2 / Val(paridadf)
End If
If moneda3 = "S" Then
   xpago3 = xpago3 / Val(paridadf)
End If
End If

Label20 = "Falta"
sdx = xpago1 + xpago2 + xpago3
vuelto = Val(total) - Format(sdx, "0.00")
If Val(vuelto) <= 0 Then
   Label20 = "Vuelto"
   Exit Sub
End If
End Sub

Private Sub nombre_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
End Sub

Private Sub fpago_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   KeyAscii = 0
   Exit Sub
End If
suma_fpago
pago1.SetFocus
End Sub

Private Sub fpago_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_fpago 3
End If
End Sub

Private Sub fpago2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   KeyAscii = 0
   Exit Sub
End If
suma_fpago
pago2.SetFocus

End Sub

Private Sub fpago2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_fpago 7
End If

End Sub

Private Sub fpago3_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   KeyAscii = 0
   Exit Sub
End If
suma_fpago
pago3.SetFocus

End Sub

Private Sub fpago3_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_fpago 8
End If

End Sub

Private Sub Label5_Click()
tficha.Show 1
End Sub

Private Sub Label6_Click()
tficha.Show 1

End Sub

Private Sub moneda_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
paridad.SetFocus

End Sub

Private Sub numero_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
tipocliente.SetFocus

End Sub

Private Sub observa_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
DBGrid3.Refresh
               DBGrid3.Col = 0
               DBGrid3.Row = DBGrid3.VisibleRows - 1
               DBGrid3.SetFocus

End Sub

Private Sub pago1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
suma_fpago
fpago2.SetFocus
End Sub

Private Sub pago2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
suma_fpago
fpago3.SetFocus
End Sub

Private Sub pago3_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
suma_fpago
End Sub

Private Sub paridad_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
cajero.SetFocus

End Sub

Private Sub ruc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
moneda.SetFocus
End Sub

Private Sub serie_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Len(serie) = 0 Then
   serie.SetFocus
   Exit Sub
End If
numero.SetFocus

End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
bodega.SetFocus

End Sub

Private Sub tipo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Len(tipo) = 0 Then
   tipo.SetFocus
   Exit Sub
End If
If existe_tipo() = 0 Then
   MsgBox "Tipo Documento no existe", 48, "Aviso"
   tipo.SetFocus
   Exit Sub
End If
serie.SetFocus
End Sub

Private Sub tipo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_tipo
End If

End Sub

Private Sub tipocliente_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If tipocliente <> "P" And tipocliente <> "A" Then
   tipocliente.SetFocus
   Exit Sub
End If
codigo.SetFocus

End Sub

Private Sub transporte_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
bodega.SetFocus
End Sub

Private Sub transporte_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_transporte
End If

End Sub

Private Sub turno_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
fecha.SetFocus

End Sub

Private Sub vendedor_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   KeyAscii = 0
   Exit Sub
End If
transporte.SetFocus
End Sub

Private Sub vendedor_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_vendedor
End If

End Sub
Sub consulta_tipo()
Dim cad As String
   Dim rconsulta As New ADODB.Recordset
   cad = "SELECT * FROM TIPODOC  "
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True Or rconsulta.BOF = True Then
      Exit Sub
   End If

Frame2.Visible = True
Frame2.Enabled = True
buffer = ""
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Tipodoc"
Combo1.ListIndex = 0
opcion1 = 1
buffer.SetFocus
Command1_Click
End Sub

Sub consulta_codigo()
Dim cad As String
   Dim rconsulta As New ADODB.Recordset
   cad = "SELECT * FROM Cliente  "
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True Or rconsulta.BOF = True Then
      Exit Sub
   End If
Frame2.Visible = True
Frame2.Enabled = True
buffer = ""
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Cliente"
Combo1.ListIndex = 0
opcion1 = 2
buffer.SetFocus
Command1_Click
End Sub
Sub consulta_fpago(sw As Integer)
Dim cad As String
   Dim rconsulta As New ADODB.Recordset
   cad = "SELECT fpago FROM fpago  "
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True Or rconsulta.BOF = True Then
      Exit Sub
   End If

Frame2.Visible = True
Frame2.Enabled = True
buffer = ""
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Fpago"
Combo1.ListIndex = 0
If sw = 3 Then
   opcion1 = 3
End If
If sw = 7 Then
   opcion1 = 7
End If
If sw = 8 Then
   opcion1 = 8
End If
buffer.SetFocus
Command1_Click
End Sub
Sub consulta_cajero()
Dim cad As String
   Dim rconsulta As New ADODB.Recordset
   cad = "SELECT  * FROM personal  "
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True Or rconsulta.BOF = True Then
      Exit Sub
   End If

Frame2.Visible = True
Frame2.Enabled = True
buffer = ""
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Personal"
Combo1.ListIndex = 0
opcion1 = 30
buffer.SetFocus
Command1_Click
End Sub
Sub consulta_transporte()
Dim cad As String
   Dim rconsulta As New ADODB.Recordset
   cad = "SELECT  * FROM transporte  "
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True Or rconsulta.BOF = True Then
      Exit Sub
   End If

Frame2.Visible = True
Frame2.Enabled = True
buffer = ""
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Transporte"
Combo1.ListIndex = 0
opcion1 = 32
buffer.SetFocus
Command1_Click
End Sub
Sub consulta_bodega()
Dim cad As String
   Dim rconsulta As New ADODB.Recordset
   cad = "SELECT  * FROM bodega  "
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True Or rconsulta.BOF = True Then
      Exit Sub
   End If

Frame2.Visible = True
Frame2.Enabled = True
buffer = ""
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Bodega"
Combo1.ListIndex = 0
opcion1 = 33
buffer.SetFocus
Command1_Click
End Sub

Sub consulta_caja()
Dim cad As String
   Dim rconsulta As New ADODB.Recordset
   cad = "SELECT  * FROM caja  "
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True Or rconsulta.BOF = True Then
      Exit Sub
   End If

Frame2.Visible = True
Frame2.Enabled = True
buffer = ""
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Caja"
Combo1.ListIndex = 0
opcion1 = 31
buffer.SetFocus
Command1_Click
End Sub

Sub consulta_vendedor()
Dim cad As String
   Dim rconsulta As New ADODB.Recordset
   cad = "SELECT  * FROM personal  "
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True Or rconsulta.BOF = True Then
      Exit Sub
   End If

Frame2.Visible = True
Frame2.Enabled = True
buffer = ""
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Personal"
Combo1.ListIndex = 0
opcion1 = 4
buffer.SetFocus
Command1_Click
End Sub
Sub consulta_producto()
Dim cad As String
   Dim rconsulta As New ADODB.Recordset
   cad = "SELECT  * FROM producto  "
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True Or rconsulta.BOF = True Then
      Exit Sub
   End If

Frame2.Visible = True
Frame2.Enabled = True
buffer = ""
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Producto"
Combo1.ListIndex = 0
opcion1 = 5
buffer.SetFocus
Command1_Click
End Sub
Sub consulta_tratamiento()
Dim cad As String '
   Dim rconsulta As New ADODB.Recordset
   cad = "SELECT  * FROM tratamiento where cliente='" & Trim(codigo) & "'"
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True Or rconsulta.BOF = True Then
      Exit Sub
   End If

Frame2.Visible = True
Frame2.Enabled = True
buffer = ""
Combo1.Clear
Combo1.AddItem "Tratamiento"
Combo1.ListIndex = 0
opcion1 = 51
buffer.SetFocus
Command1_Click
End Sub




Sub ejecuta(sw As Integer)
Dim rconsulta As New ADODB.Recordset
Dim cad As String
If opcion1 = 50 Then  'consultas
   If Len(buffer) = 0 Then
      cad = "SELECT Consulta,Partisaldo,Empresaldo,Sede,Cliente FROM consulta  where sede='" & sede & "' and cliente='" & Trim(codigo) & "'"
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT Consulta,Partisaldo,Empresaldo,Sede,Cliente FROM consulta where  sede='" & Trim(sede) & "' and cliente='" & Trim(codigo) & "' and " & Combo1 & " like '" & buffer & "%'"
   End If
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True Or rconsulta.BOF = True Then
      buffer.SetFocus
      Exit Sub
   End If
   Set dbGrid2.DataSource = rconsulta
   dbGrid2.Columns(0).Width = 1000
   dbGrid2.Columns(1).Width = 1000
   dbGrid2.Columns(2).Width = 1000
   dbGrid2.Columns(3).Width = 1000
   If sw = 1 Then
      dbGrid2.SetFocus
   End If
End If
If opcion1 = 51 Then  'Tratamiento
   If Len(buffer) = 0 Then
      cad = "SELECT Tratamiento,PagaParticular,PagaEmpresa,Sede,Cliente ,Consulta FROM tratamiento  where sede='" & sede & "' and cliente='" & Trim(codigo) & "'"
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT Tratamiento,PagaParticular,PagaEmpresa,Sede,Cliente ,Consulta FROM tratamiento where  sede='" & Trim(sede) & "' and cliente='" & Trim(codigo) & "' and " & Combo1 & " like '" & buffer & "%'"
   End If
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True Or rconsulta.BOF = True Then
      buffer.SetFocus
      Exit Sub
   End If
   Set dbGrid2.DataSource = rconsulta
   dbGrid2.Columns(0).Width = 1000
   dbGrid2.Columns(1).Width = 1000
   dbGrid2.Columns(2).Width = 1000
   dbGrid2.Columns(3).Width = 1000
   If sw = 1 Then
      dbGrid2.SetFocus
   End If
End If

If opcion1 = 2 Then  'clientes
   If Len(buffer) = 0 Then
      cad = "SELECT Nombre,Cliente,Ruc FROM Cliente  "
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT Nombre,Cliente,Ruc FROM Cliente where  " & Combo1 & " like '" & buffer & "%'"
   End If
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True Or rconsulta.BOF = True Then
      buffer.SetFocus
      Exit Sub
   End If
   Set dbGrid2.DataSource = rconsulta
   dbGrid2.Columns(0).Width = 5000
   dbGrid2.Columns(1).Width = 1000
   If sw = 1 Then
      dbGrid2.SetFocus
   End If
End If
If opcion1 = 1 Then  'clientes
   If Len(buffer) = 0 Then
      cad = "SELECT Nombre,Tipodoc FROM Tipodoc  "
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT Nombre,tipodoc FROM tipodoc where  " & Combo1 & " like '" & buffer & "%'"
   End If
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      buffer.SetFocus
      Exit Sub
   End If
   Set dbGrid2.DataSource = rconsulta
   dbGrid2.Columns(0).Width = 5000
   dbGrid2.Columns(1).Width = 1000
   If sw = 1 Then
      dbGrid2.SetFocus
   End If
End If
If opcion1 = 3 Then  'clientes
If Len(buffer) = 0 Then
      cad = "SELECT Nombre,Fpago,moneda FROM Fpago  "
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT Nombre,Fpago,moneda FROM Fpago where  " & Combo1 & " like '" & buffer & "%'"
   End If
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      buffer.SetFocus
      Exit Sub
   End If
   Set dbGrid2.DataSource = rconsulta
   dbGrid2.Columns(0).Width = 5000
   dbGrid2.Columns(1).Width = 1000
   If sw = 1 Then
      dbGrid2.SetFocus
   End If
End If
If opcion1 = 7 Then  'clientes
If Len(buffer) = 0 Then
      cad = "SELECT Nombre,Fpago,moneda FROM Fpago  "
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT Nombre,Fpago,moneda FROM Fpago where  " & Combo1 & " like '" & buffer & "%'"
   End If
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      buffer.SetFocus
      Exit Sub
   End If
   Set dbGrid2.DataSource = rconsulta
   dbGrid2.Columns(0).Width = 5000
   dbGrid2.Columns(1).Width = 1000
   If sw = 1 Then
      dbGrid2.SetFocus
   End If
End If
If opcion1 = 8 Then  'clientes
If Len(buffer) = 0 Then
      cad = "SELECT Nombre,Fpago,moneda FROM Fpago  "
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT Nombre,Fpago,moneda FROM Fpago where  " & Combo1 & " like '" & buffer & "%'"
   End If
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      buffer.SetFocus
      Exit Sub
   End If
   Set dbGrid2.DataSource = rconsulta
   dbGrid2.Columns(0).Width = 5000
   dbGrid2.Columns(1).Width = 1000
   If sw = 1 Then
      dbGrid2.SetFocus
   End If
End If



If opcion1 = 4 Then  'clientes
If Len(buffer) = 0 Then
      cad = "SELECT Nombre,Personal FROM Personal  "
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT Nombre,Personal FROM Personal where  " & Combo1 & " like '" & buffer & "%'"
   End If
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      buffer.SetFocus
      Exit Sub
   End If
   Set dbGrid2.DataSource = rconsulta
   dbGrid2.Columns(0).Width = 5000
   dbGrid2.Columns(1).Width = 1000
   If sw = 1 Then
      dbGrid2.SetFocus
   End If
End If
If opcion1 = 30 Then  'cajero
If Len(buffer) = 0 Then
      cad = "SELECT Nombre,Personal FROM Personal  "
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT Nombre,Personal FROM Personal where  " & Combo1 & " like '" & buffer & "%'"
   End If
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      buffer.SetFocus
      Exit Sub
   End If
   Set dbGrid2.DataSource = rconsulta
   dbGrid2.Columns(0).Width = 5000
   dbGrid2.Columns(1).Width = 1000
   If sw = 1 Then
      dbGrid2.SetFocus
   End If
End If
If opcion1 = 31 Then  'cajero
If Len(buffer) = 0 Then
      cad = "SELECT Nombre,Caja FROM caja where sede='" & Trim(sede) & "'"
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT Nombre,Caja FROM Caja where sede='" & Trim(sede) & "' and " & Combo1 & " like '" & buffer & "%'"
   End If
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      buffer.SetFocus
      Exit Sub
   End If
   Set dbGrid2.DataSource = rconsulta
   dbGrid2.Columns(0).Width = 5000
   dbGrid2.Columns(1).Width = 1000
   If sw = 1 Then
      dbGrid2.SetFocus
   End If
End If
If opcion1 = 32 Then  'cajero
If Len(buffer) = 0 Then
      cad = "SELECT Nombre,Transporte FROM Transporte  "
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT Nombre,Transporte FROM Transporte where  " & Combo1 & " like '" & buffer & "%'"
   End If
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      buffer.SetFocus
      Exit Sub
   End If
   Set dbGrid2.DataSource = rconsulta
   dbGrid2.Columns(0).Width = 5000
   dbGrid2.Columns(1).Width = 1000
   If sw = 1 Then
      dbGrid2.SetFocus
   End If
End If
If opcion1 = 33 Then  'cajero
If Len(buffer) = 0 Then
      cad = "SELECT Nombre,Bodega FROM Bodega  where sede='" & sede & "'"
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT Nombre,Bodega FROM Bodega where sede='" & sede & "' and " & Combo1 & " like '" & buffer & "%'"
   End If
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      buffer.SetFocus
      Exit Sub
   End If
   Set dbGrid2.DataSource = rconsulta
   dbGrid2.Columns(0).Width = 5000
   dbGrid2.Columns(1).Width = 1000
   If sw = 1 Then
      dbGrid2.SetFocus
   End If
End If
If opcion1 = 5 Then  'clientes
If Len(buffer) = 0 Then
      cad = "SELECT Nombre,Producto,Precio,Igv FROM producto  "
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT Nombre,producto,Precio,Igv FROM producto where  " & Combo1 & " like '" & buffer & "%'"
   End If
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      buffer.SetFocus
      Exit Sub
   End If
   Set dbGrid2.DataSource = rconsulta
   dbGrid2.Columns(0).Width = 5000
   dbGrid2.Columns(1).Width = 1000
   dbGrid2.Columns(2).Width = 1000
   dbGrid2.Columns(3).Width = 1000
   If sw = 1 Then
      dbGrid2.SetFocus
   End If
End If


End Sub
Function existe_numero()
Dim rs1 As New ADODB.Recordset
existe_numero = 1
If rs1.State = 1 Then rs1.Close
   rs1.Open "SELECT sede,tipo,serie,Numero FROM factura where sede='" & sede & "' and caja='" & Trim(caja) & "' and tipo='" & Trim(tipo) & "' and serie='" & Trim(serie) & "' and numero='" & Trim(numero) & "'", cn, adOpenDynamic, adLockReadOnly
   If rs1.EOF Then
      existe_numero = 0
   End If
   rs1.Close
   Set rs1 = Nothing
   
   
End Function
Function existe_codigo()
Dim rs1 As New ADODB.Recordset
   existe_codigo = 1
   If rs1.State = 1 Then rs1.Close
   rs1.Open "SELECT cliente FROM cliente where cliente='" & codigo & "'", cn, adOpenDynamic, adLockReadOnly
   If rs1.EOF Then
      existe_codigo = 0
   End If
   rs1.Close
   Set rs1 = Nothing
End Function
Function existe_tipo()
Dim rs1 As New ADODB.Recordset
existe_tipo = 1
If rs1.State = 1 Then rs1.Close
   rs1.Open "SELECT tipodoc FROM tipodoc where tipodoc='" & tipo & "'", cn, adOpenDynamic, adLockReadOnly
   If rs1.EOF Then
      existe_tipo = 0
   End If
   rs1.Close
   Set rs1 = Nothing

End Function
Sub carga_detalle()
Dim buf As String
Dim found As Integer
On Error GoTo cmd34_err
buf = "select * from " & dgusuario
               Data1.Connect = "foxpro 2.5;"
               Data1.DatabaseName = globaldat
               Data1.RecordSource = buf
               Data1.Refresh
               DBGrid3.Refresh
               found = sumar_detalle()
               DBGrid3.Col = 0
               DBGrid3.Row = DBGrid3.VisibleRows - 1
               DBGrid3.SetFocus
Exit Sub
cmd34_err:
MsgBox "Error en select " & Error$, 48, "Aviso"
Exit Sub
End Sub
Function copiar_tmp(buf As String)
On Error GoTo cmd35_err
FileCopy globaldat & "\tmpdeta.dbf", globaldat & "\" & buf & ".dbf"
copiar_tmp = 1
Exit Function
cmd35_err:
MsgBox " Error al Copiar Tmp ", 48, "Aviso"
Exit Function
End Function
Function sumar_detalle()
On Error GoTo cmd48_err
Dim xcantidad As Double
Dim xtotal As Double
Dim xneto As Double
Dim xdescuento As Double
Dim xsubtotal As Double
Dim ximpuesto As Double
xcantidad = 0
xtotal = 0
xneto = 0
xdescuento = 0
xsubtotal = 0
ximpuesto = 0
cantidad = Format(xcantidad, "0.00")
neto = Format(xneto, "0.00")
descuento = Format(xdescuento, "0.00")
subtotal = Format(xsubtotal, "0.00")
impuesto = Format(ximpuesto, "0.00")
total = Format(xtotal, "0.00")

inicio_data1
Do
If Data1.Recordset.EOF Then Exit Do
xcantidad = xcantidad + Val("" & Data1.Recordset.Fields("cantidad"))
xneto = xneto + Val("" & Data1.Recordset.Fields("neto"))
xdescuento = xdescuento + Val("" & Data1.Recordset.Fields("descuento"))
xsubtotal = xsubtotal + Val("" & Data1.Recordset.Fields("subtotal"))
ximpuesto = ximpuesto + Val("" & Data1.Recordset.Fields("impuesto"))
xtotal = xtotal + Val("" & Data1.Recordset.Fields("total"))
Data1.Recordset.MoveNext
Loop
cantidad = Format(xcantidad, "0.00")
neto = Format(xneto, "0.00")
descuento = Format(xdescuento, "0.00")
subtotal = Format(xsubtotal, "0.00")
impuesto = Format(ximpuesto, "0.00")
total = Format(xtotal, "0.00")
Exit Function
cmd48_err:
'MsgBox "Aviso en sumar detalle " + Error$, 48, "Aviso"
Exit Function


End Function
Sub calcula_datos()
Dim xcantidad As Double
Dim xtotal As Double
Dim xneto As Double
Dim xdescuento As Double
Dim xsubtotal As Double
Dim ximpuesto As Double
Dim xx As Double
xcantidad = 0
xtotal = 0
xneto = 0
xdescuento = 0
xsubtotal = 0
ximpuesto = 0

xneto = Val("" & DBGrid3.Columns(4)) * Val("" & DBGrid3.Columns(5))
xdescuento = xneto * Val("" & DBGrid3.Columns(6)) / 100
xtotal = xneto - xdescuento
          xx = Val("" & DBGrid3.Columns(12))
          xsubtotal = xtotal / (1 + (xx) / 100) 'subtotal + isc
          ximpuesto = xtotal - xsubtotal
          
DBGrid3.Columns(11) = xneto
DBGrid3.Columns(10) = xdescuento
DBGrid3.Columns(9) = xsubtotal
DBGrid3.Columns(8) = ximpuesto
DBGrid3.Columns(7) = xtotal

End Sub
Sub cerrar_data1()
On Error GoTo cmd451_err
Data1.Recordset.Close
Exit Sub
cmd451_err:
Exit Sub
End Sub
Sub inicio_data1()
On Error GoTo cmd451_err
Data1.Recordset.MoveFirst
Exit Sub
cmd451_err:
Exit Sub
End Sub

Sub grabar_detalle()
Dim cad As String
inicio_data1
Do
If Data1.Recordset.EOF Then Exit Do
   cad = "INSERT INTO detalle VALUES('" & Trim(sede) & "','" & Trim(tipo) & "','" & Trim(serie) & "','" & Trim(numero) & "','" & Trim("" & Data1.Recordset.Fields("producto")) & "','" & Trim("" & Data1.Recordset.Fields("descripcio")) & "'," & Val("" & Data1.Recordset.Fields("unidad")) & "," & Val("" & Data1.Recordset.Fields("factor")) & "," & Val("" & Data1.Recordset.Fields("cantidad")) & "," & Val("" & Data1.Recordset.Fields("precio")) & "," & Val("" & Data1.Recordset.Fields("total")) & "," & Val("" & Data1.Recordset.Fields("impuesto")) & "," & Val("" & Data1.Recordset.Fields("subtotal")) & "," & Val("" & Data1.Recordset.Fields("descuento")) & "," & Val("" & Data1.Recordset.Fields("neto")) & "," & Val("" & Data1.Recordset.Fields("igv")) & "," & Val("" & Data1.Recordset.Fields("dscto")) & ",'" & Trim("" & Data1.Recordset.Fields("tratamient")) & "','" & Trim("" & Data1.Recordset.Fields("consulta")) & "','" & Trim("" & Data1.Recordset.Fields("sedecotra")) & "'"
   cad = cad & ",'" & Trim("" & Data1.Recordset.Fields("tipocobro")) & "'"
   cad = cad & ",'" & Trim(cajero) & "','" & Trim(caja) & "','" & Trim(turno) & "','" & Trim(bodega) & "')"
   cn.Execute (cad)
   Data1.Recordset.MoveNext
Loop
End Sub
Sub carga_detalle_datos()
On Error GoTo cmd9_err
Dim cad As String
Dim rss As New ADODB.Recordset
Data1.Database.Execute "DELETE FROM " & dgusuario
Data1.Refresh

cad = "SELECT * FROM detalle where sede='" & sede & "' and tipo='" & tipo & "' and serie='" & serie & "' and numero='" & numero & "'"
If rss.State = 1 Then rss.Close
rss.Open cad, cn, adOpenStatic, adLockOptimistic
Do
 If rss.EOF Then Exit Do
    Data1.Recordset.AddNew
    Data1.Recordset.Fields("producto") = "" & rss.Fields("producto")
    Data1.Recordset.Fields("DESCRIPCIO") = "" & rss.Fields("descripcio")
    Data1.Recordset.Fields("unidad") = "" & rss.Fields("unidad")
    Data1.Recordset.Fields("factor") = Val("" & rss.Fields("factor"))
    Data1.Recordset.Fields("precio") = Val("" & rss.Fields("precio"))
    Data1.Recordset.Fields("dscto") = Val("" & rss.Fields("dscto"))
    Data1.Recordset.Fields("cantidad") = Val("" & rss.Fields("cantidad"))
    Data1.Recordset.Fields("total") = Val("" & rss.Fields("total"))
    Data1.Recordset.Fields("impuesto") = Val("" & rss.Fields("impuesto"))
    Data1.Recordset.Fields("subtotal") = Val("" & rss.Fields("subtotal"))
    Data1.Recordset.Fields("descuento") = Val("" & rss.Fields("descuento"))
    Data1.Recordset.Fields("neto") = Val("" & rss.Fields("neto"))
    Data1.Recordset.Fields("igv") = Val("" & rss.Fields("igv"))
    Data1.Recordset.Fields("consulta") = "" & rss.Fields("consulta")
    Data1.Recordset.Fields("tratamient") = "" & rss.Fields("tratamiento")
    Data1.Recordset.Fields("sedecotra") = "" & rss.Fields("sedecotra")
    Data1.Recordset.Fields("tipocobro") = "" & rss.Fields("tipocobro")
    Data1.Recordset.Update
    rss.MoveNext
Loop
rss.Close
Set rss = Nothing
Exit Sub
cmd9_err:
MsgBox "Aviso en carga detalle datos " + Error$, 48, "Aviso"
Exit Sub
End Sub
Sub carga_detalle_borra()
Data1.Database.Execute "DELETE FROM " & dgusuario
Data1.Refresh
End Sub
Function fpago_nombre(buf As String) As String
Dim rs1 As New ADODB.Recordset
If rs1.State = 1 Then rs1.Close
   rs1.Open "SELECT Nombre FROM fpago where fpago='" & buf & "'", cn, adOpenDynamic, adLockReadOnly
   If Not rs1.EOF Then
      fpago_nombre = "" & rs1.Fields("nombre")
   End If
   rs1.Close
   Set rs1 = Nothing
   
   
End Function

Function busca_paridadf() As Double
Dim rs1 As New ADODB.Recordset
paridadf = 0
If rs1.State = 1 Then rs1.Close
   rs1.Open "SELECT compra FROM moneda  where moneda='D'", cn, adOpenDynamic, adLockReadOnly
   If Not rs1.EOF Then
      paridadf = Val("" & rs1.Fields("compra"))
   End If
   rs1.Close
   Set rs1 = Nothing
   
   
End Function
Function consulta_saldo()
On Error GoTo cmd99_err
Dim cad As String
Dim rss As New ADODB.Recordset
Dim sdx As Double
Dim sdx1 As Double
sdx = 0
xsaldo = ""
cad = "SELECT parsaldo FROM tratamiento where  cliente='" & codigo & "'"
If rss.State = 1 Then rss.Close
rss.Open cad, cn, adOpenStatic, adLockOptimistic
Do
 If rss.EOF Then Exit Do
    sdx = sdx + Val("" & rss.Fields("parsaldo"))
    rss.MoveNext
Loop
rss.Close
xsaldo = Format(sdx, "0.00")
Set rss = Nothing
Exit Function
cmd99_err:
MsgBox "Aviso en consulta saldo " + Error$, 48, "Aviso"
Exit Function

End Function
Function codigo_nombre(buf As String) As String
Dim rs1 As New ADODB.Recordset
If rs1.State = 1 Then rs1.Close
   rs1.Open "SELECT Nombre FROM cliente where cliente='" & buf & "'", cn, adOpenDynamic, adLockReadOnly
   If Not rs1.EOF Then
      codigo_nombre = "" & rs1.Fields("nombre")
   End If
   rs1.Close
   Set rs1 = Nothing
   
   
End Function
Function existe_cajero(buf As String)
Dim rs1 As New ADODB.Recordset
existe_cajero = 1
If rs1.State = 1 Then rs1.Close
   rs1.Open "SELECT personal FROM personal where personal='" & buf & "'", cn, adOpenDynamic, adLockReadOnly
   If rs1.EOF Then
      existe_cajero = 0
   End If
   rs1.Close
   Set rs1 = Nothing

End Function
Function existe_caja(buf As String)
Dim rs1 As New ADODB.Recordset
existe_caja = 1
If rs1.State = 1 Then rs1.Close
   rs1.Open "SELECT caja FROM caja where caja='" & buf & "'", cn, adOpenDynamic, adLockReadOnly
   If rs1.EOF Then
      existe_caja = 0
   End If
   rs1.Close
   Set rs1 = Nothing

End Function
Function existe_turno(buf As String)
Dim rs1 As New ADODB.Recordset
existe_turno = 1
If rs1.State = 1 Then rs1.Close
   rs1.Open "SELECT turno FROM turno where turno='" & buf & "'", cn, adOpenDynamic, adLockReadOnly
   If rs1.EOF Then
      existe_turno = 0
   End If
   rs1.Close
   Set rs1 = Nothing

End Function
Function existe_transporte(buf As String)
Dim rs1 As New ADODB.Recordset
existe_transporte = 1
If rs1.State = 1 Then rs1.Close
   rs1.Open "SELECT transporte FROM transporte where transporte='" & buf & "'", cn, adOpenDynamic, adLockReadOnly
   If rs1.EOF Then
      existe_transporte = 0
   End If
   rs1.Close
   Set rs1 = Nothing

End Function
Function existe_bodega(buf As String)
Dim rs1 As New ADODB.Recordset
existe_bodega = 1
If rs1.State = 1 Then rs1.Close
   rs1.Open "SELECT bodega FROM bodega where sede='" & Trim(sede) & "' and bodega='" & buf & "'", cn, adOpenDynamic, adLockReadOnly
   If rs1.EOF Then
      existe_bodega = 0
   End If
   rs1.Close
   Set rs1 = Nothing

End Function
Function existe_sede(buf As String)
Dim rs1 As New ADODB.Recordset
existe_sede = 1
If rs1.State = 1 Then rs1.Close
   rs1.Open "SELECT sede FROM sede where sede='" & Trim(buf) & "'", cn, adOpenDynamic, adLockReadOnly
   If rs1.EOF Then
      existe_sede = 0
   End If
   rs1.Close
   Set rs1 = Nothing

End Function









Private Sub xcliente_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   'consulta_xcliente
End If
End Sub
