VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form treceta 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recetas de productos"
   ClientHeight    =   8685
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consulta"
      Height          =   8415
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Visible         =   0   'False
      Width           =   12930
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
         Left            =   6435
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   570
         Width           =   1935
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
         MaxLength       =   30
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   600
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   600
         Width           =   2295
      End
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   6735
         Left            =   240
         TabIndex        =   70
         Top             =   1200
         Width           =   10920
         _ExtentX        =   19262
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
      Begin ChamaleonButton.ChameleonBtn Command1 
         Height          =   495
         Left            =   9060
         TabIndex        =   73
         Top             =   510
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   873
         BTYPE           =   5
         TX              =   "Buscar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   8421504
         FCOL            =   4210752
         FCOLO           =   4210752
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "treceta.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn ChaAceptar 
         Height          =   930
         Left            =   11220
         TabIndex        =   74
         Top             =   1230
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1640
         BTYPE           =   5
         TX              =   "Aceptar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   255
         BCOLO           =   255
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "treceta.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn ChaCerrar 
         Height          =   750
         Left            =   11295
         TabIndex        =   75
         Top             =   2430
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1323
         BTYPE           =   5
         TX              =   "Cerrar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   8421504
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "treceta.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblFamilia 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Familia"
         Height          =   495
         Left            =   5520
         TabIndex        =   79
         Top             =   585
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Captura de Insumos"
      Height          =   7965
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   8040
      Begin VB.TextBox platosi 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   6840
         TabIndex        =   82
         Text            =   "platos"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox xtotal 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2145
         MaxLength       =   10
         TabIndex        =   76
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton cmdBuscarProducto 
         Height          =   420
         Left            =   4560
         Picture         =   "treceta.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Buscar producto"
         Top             =   1635
         Width           =   435
      End
      Begin VB.TextBox t13 
         Height          =   375
         Left            =   7665
         MaxLength       =   15
         TabIndex        =   37
         Top             =   5535
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox t14 
         Height          =   375
         Left            =   7665
         MaxLength       =   15
         TabIndex        =   36
         Top             =   5895
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox t15 
         Height          =   375
         Left            =   7665
         MaxLength       =   15
         TabIndex        =   35
         Top             =   6255
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox t16 
         Height          =   375
         Left            =   7665
         MaxLength       =   15
         TabIndex        =   34
         Top             =   6615
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command2X 
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
         Height          =   810
         Left            =   2790
         Picture         =   "treceta.frx":0802
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Nuevo registro"
         Top             =   480
         Width           =   1245
      End
      Begin VB.TextBox t12 
         Height          =   375
         Left            =   5505
         MaxLength       =   15
         TabIndex        =   32
         Top             =   6615
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox t11 
         Height          =   375
         Left            =   5505
         MaxLength       =   15
         TabIndex        =   31
         Top             =   6255
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox t10 
         Height          =   375
         Left            =   5505
         MaxLength       =   15
         TabIndex        =   30
         Top             =   5895
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox t9 
         Height          =   375
         Left            =   5505
         MaxLength       =   15
         TabIndex        =   29
         Top             =   5535
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox t8 
         Height          =   375
         Left            =   3345
         MaxLength       =   15
         TabIndex        =   28
         Top             =   6615
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox t7 
         Height          =   375
         Left            =   3345
         MaxLength       =   15
         TabIndex        =   27
         Top             =   6255
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox t6 
         Height          =   375
         Left            =   3345
         MaxLength       =   15
         TabIndex        =   26
         Top             =   5895
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox t5 
         Height          =   375
         Left            =   3345
         MaxLength       =   15
         TabIndex        =   25
         Top             =   5535
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox t4 
         Height          =   375
         Left            =   1185
         MaxLength       =   15
         TabIndex        =   24
         Top             =   6615
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox t3 
         Height          =   375
         Left            =   1185
         MaxLength       =   15
         TabIndex        =   23
         Top             =   6255
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox t2 
         Height          =   375
         Left            =   1185
         MaxLength       =   15
         TabIndex        =   22
         Top             =   5895
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox t1 
         Height          =   375
         Left            =   1185
         MaxLength       =   15
         TabIndex        =   21
         Top             =   5535
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox precio 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2145
         MaxLength       =   10
         TabIndex        =   20
         Top             =   4020
         Width           =   1335
      End
      Begin VB.TextBox cantidad 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2145
         MaxLength       =   10
         TabIndex        =   19
         Top             =   3675
         Width           =   1335
      End
      Begin VB.TextBox productoi 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2145
         MaxLength       =   15
         TabIndex        =   18
         Top             =   1665
         Width           =   2295
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
         Height          =   810
         Left            =   585
         MaskColor       =   &H00E0E0E0&
         Picture         =   "treceta.frx":1A14
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Grabar registro"
         Top             =   480
         Width           =   1245
      End
      Begin VB.CommandButton Command1X 
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
         Height          =   810
         Left            =   6480
         MaskColor       =   &H00E0E0E0&
         Picture         =   "treceta.frx":2C26
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Salir"
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label platos2 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3600
         TabIndex        =   83
         Top             =   3720
         Width           =   1560
      End
      Begin VB.Line Line1 
         X1              =   330
         X2              =   7800
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total >>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   585
         TabIndex        =   77
         Top             =   4500
         Width           =   990
      End
      Begin VB.Label nt13 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6585
         TabIndex        =   66
         Top             =   5535
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label nt14 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6585
         TabIndex        =   65
         Top             =   5895
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label nt15 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6585
         TabIndex        =   64
         Top             =   6255
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label nt16 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6585
         TabIndex        =   63
         Top             =   6615
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label nlinea 
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   4065
         TabIndex        =   62
         Top             =   5910
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.Label nt12 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4425
         TabIndex        =   61
         Top             =   6615
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label nt11 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4425
         TabIndex        =   60
         Top             =   6255
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label nt10 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4425
         TabIndex        =   59
         Top             =   5895
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label nt9 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4365
         TabIndex        =   58
         Top             =   5565
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label nt8 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2265
         TabIndex        =   57
         Top             =   6615
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label nt7 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2265
         TabIndex        =   56
         Top             =   6255
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label nt6 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2265
         TabIndex        =   55
         Top             =   5895
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label nt5 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2265
         TabIndex        =   54
         Top             =   5535
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label nt4 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   105
         TabIndex        =   53
         Top             =   6615
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label nt3 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   105
         TabIndex        =   52
         Top             =   6255
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tallas"
         Height          =   375
         Left            =   6870
         TabIndex        =   51
         Top             =   5385
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label nt2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   105
         TabIndex        =   50
         Top             =   5895
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label nt1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   105
         TabIndex        =   49
         Top             =   5535
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   585
         TabIndex        =   48
         Top             =   4020
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   585
         TabIndex        =   47
         Top             =   3660
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Linea"
         Height          =   375
         Left            =   7350
         TabIndex        =   46
         Top             =   6195
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label factor 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2145
         TabIndex        =   45
         Top             =   3030
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Factor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   585
         TabIndex        =   44
         Top             =   3015
         Width           =   1575
      End
      Begin VB.Label unidad 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2145
         TabIndex        =   43
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   585
         TabIndex        =   42
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label descripcioi 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   41
         Top             =   2235
         Width           =   5625
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   585
         TabIndex        =   40
         Top             =   2250
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   585
         TabIndex        =   39
         Top             =   1665
         Width           =   1575
      End
      Begin VB.Label lineai 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   7410
         TabIndex        =   38
         Top             =   5865
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.TextBox detalle 
      Height          =   5175
      Left            =   8160
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   2760
      Width           =   4575
   End
   Begin VB.TextBox nro 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   2
      Text            =   "1"
      Top             =   720
      Width           =   975
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
      Left            =   1080
      MaskColor       =   &H00E0E0E0&
      Picture         =   "treceta.frx":3E38
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   1095
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
      Picture         =   "treceta.frx":504A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo registro"
      Top             =   0
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid dbgrid2 
      Height          =   5175
      Left            =   0
      TabIndex        =   3
      Top             =   2760
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   22
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   24
      BeginProperty Column00 
         DataField       =   "Productoi"
         Caption         =   "Producto"
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
         DataField       =   "Descripcio"
         Caption         =   "Descripcio"
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
         DataField       =   "Unidad"
         Caption         =   "Und"
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
         DataField       =   "Factor"
         Caption         =   "Fx"
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
         DataField       =   "Cantidad"
         Caption         =   "Cant"
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
         DataField       =   "Linea"
         Caption         =   "Linea"
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
      BeginProperty Column08 
         DataField       =   "t1"
         Caption         =   "t1"
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
      BeginProperty Column09 
         DataField       =   "t2"
         Caption         =   "t2"
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
      BeginProperty Column10 
         DataField       =   "t3"
         Caption         =   "t3"
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
      BeginProperty Column11 
         DataField       =   "t4"
         Caption         =   "t4"
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
      BeginProperty Column12 
         DataField       =   "t5"
         Caption         =   "t5"
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
      BeginProperty Column13 
         DataField       =   "t6"
         Caption         =   "t6"
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
      BeginProperty Column14 
         DataField       =   "t7"
         Caption         =   "t7"
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
      BeginProperty Column15 
         DataField       =   "t8"
         Caption         =   "t8"
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
      BeginProperty Column16 
         DataField       =   "t9"
         Caption         =   "t9"
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
      BeginProperty Column17 
         DataField       =   "t10"
         Caption         =   "t10"
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
      BeginProperty Column18 
         DataField       =   "t11"
         Caption         =   "t11"
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
      BeginProperty Column19 
         DataField       =   "t12"
         Caption         =   "t12"
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
      BeginProperty Column20 
         DataField       =   "t13"
         Caption         =   "t13"
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
      BeginProperty Column21 
         DataField       =   "t14"
         Caption         =   "t14"
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
      BeginProperty Column22 
         DataField       =   "t15"
         Caption         =   "t15"
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
      BeginProperty Column23 
         DataField       =   "t16"
         Caption         =   "t16"
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
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3509.858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
         EndProperty
         BeginProperty Column13 
         EndProperty
         BeginProperty Column14 
         EndProperty
         BeginProperty Column15 
         EndProperty
         BeginProperty Column16 
         EndProperty
         BeginProperty Column17 
         EndProperty
         BeginProperty Column18 
         EndProperty
         BeginProperty Column19 
         EndProperty
         BeginProperty Column20 
         EndProperty
         BeginProperty Column21 
         EndProperty
         BeginProperty Column22 
         EndProperty
         BeginProperty Column23 
         EndProperty
      EndProperty
   End
   Begin ChamaleonButton.ChameleonBtn Label9 
      Height          =   465
      Left            =   11550
      TabIndex        =   80
      Top             =   2295
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   820
      BTYPE           =   4
      TX              =   "Grabar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "treceta.frx":625C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblProducto 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Modo de Preparación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   8175
      TabIndex        =   81
      Top             =   2295
      Width           =   3375
   End
   Begin VB.Label tiporeceta 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   9000
      TabIndex        =   71
      Top             =   120
      Width           =   45
   End
   Begin VB.Label platos 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label total 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6240
      TabIndex        =   12
      Top             =   8040
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nro Receta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label producto 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label descripcio 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   1800
      Width           =   6480
   End
   Begin VB.Label linea 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cod Barras"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Receta Para"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Menu dki232 
      Caption         =   "&Nuevo"
   End
   Begin VB.Menu dmi33 
      Caption         =   "&Modifica"
   End
   Begin VB.Menu fd2b34 
      Caption         =   "&Borra"
   End
   Begin VB.Menu clko923 
      Caption         =   "&Imprime"
   End
   Begin VB.Menu lfdo33 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "treceta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mytablev     As New ADODB.Recordset

Dim tipodereceta As String

Private Sub buffer_Change()
    ejecuta 1

End Sub

Private Sub buffer_DblClick()
    Command1_Click

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        lfdo33_Click
        Exit Sub

    End If

    Command1_Click

End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    precio.SetFocus

    ''' 11/12/2017 SubReceta
    'xtotal = Format(Val(cantidad) * Val(precio), "0.00")
    xtotal = Format(Val(cantidad) * Val(precio), "0.0000")

    ''' 11/12/2017 SubReceta
End Sub

Private Sub Label9_Click()

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd7777_err

    If Len(Trim(detalle)) > 0 Then
        cn.Execute ("update producto set detalle='" & detalle & "' where producto='" & Trim(producto) & "'")

    End If

    MsgBox "Proceso Grabado ", 48, "Aviso"
    Exit Sub
cmd7777_err:
    MsgBox "Aviso en grabar Preparacion " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub ChaAceptar_Click()

    Dim found    As Integer

    Dim rsexiste As New ADODB.Recordset

    ' AQUI YA
    '13/08/2018 Integración FE - Pizzeria
    'Cambios Pizzeria 24/05/2018
    ' rsexiste.Open "SELECT * FROM " & tiporeceta & " where  nro='" & Trim(nro) & "' and producto='" & Trim(producto) & "' and productoi='" & Trim(dbGrid1.columns(1)) & "'", cn, adOpenKeyset, adLockOptimistic
    '   If rsexiste.RecordCount > 0 Then  'si existe
    '      MsgBox "Ya existe insumo ", 48, "Aviso"
    '      Exit Sub
    '   End If
    'Cambios Pizzeria 24/05/2018

    found = busca_producto("" & dbGrid1.columns(1))
    'productoi = dbGrid1.columns(1)
    'descripcioi = dbGrid1.columns(0)
    'unidad = dbGrid1.columns(2)
    'factor = dbGrid1.columns(3)
    'precio = dbGrid1.columns(4)
    cantidad = "1"
   
    ''' 11/12/2017 SubReceta
   
    If platos > 1 Then
        platos2.Visible = True
        platos2.Caption = "(Para " & platos & " Und.) "
     
        cantidad = Format(Val(cantidad) / Val(platos), "0.0000")
    Else
        platos2.Visible = False

    End If
   
    ''' 11/12/2017 SubReceta
   
    'lineai = dbGrid1.columns(5)
    'found = busca_linea("" & lineai)
    Frame1.Visible = False
    Frame1.Enabled = False
    'xtotal = Format(Val(cantidad) * Val(precio), "0.00")
    cantidad.SetFocus

    If Len(lineai) > 0 Then
        t1.SetFocus

    End If
   
    ''' 11/12/2017 SubReceta
    'xtotal = Format(Val(cantidad) * Val(precio), "0.00")
    xtotal = Format(Val(cantidad) * Val(precio), "0.0000")

    ''' 11/12/2017 SubReceta
    '13/08/2018 Integración FE - Pizzeria
End Sub

Private Sub ChaCERRAR_Click()
    Frame1.Visible = False
    productoi.SetFocus

End Sub

Private Sub clko923_Click()

    Dim found      As Integer

    Dim I          As Integer

    Dim v          As Long

    Dim R          As Long

    Dim ih         As Integer

    Dim h          As Integer

    Dim cad        As String

    Dim Tmp        As String

    Dim sw         As Integer

    Dim mytablex   As New ADODB.Recordset

    Dim mytabley   As New ADODB.Recordset

    Dim Heading(8) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd5612_err
    
    Heading(1) = "Producto"
    Heading(2) = "Descripcio"
    Heading(3) = "Und"
    Heading(4) = "Factor"
    Heading(5) = "Cantidad"
    Heading(6) = "Costo"
    Heading(7) = "Total"
    
    mytablex.Open "Select * from producto where producto='" & producto & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then Exit Sub
         
    mytabley.Open "Select * from " & tiporeceta & " where producto='" & producto & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount = 0 Then
        mytabley.Close
        mytablex.Close
        Exit Sub

    End If
         
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_Excel(7, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
    objExcel.ActiveSheet.Cells(2, 1) = "" & mytablex.Fields("Producto")
    objExcel.ActiveSheet.Cells(2, 2) = "" & mytablex.Fields("Descripcio")
 
    objExcel.ActiveSheet.Cells(2, 3) = "Porciones"
    objExcel.ActiveSheet.Cells(2, 4) = "'(" & platos & ")"
   
    v = 4
    h = 1
    mytabley.MoveFirst
    sw = 0
    Do

        If mytabley.EOF Then Exit Do
     
        objExcel.ActiveSheet.Cells(v, h + 0) = "" & mytabley.Fields("Productoi")
        objExcel.ActiveSheet.Cells(v, h + 1) = "" & mytabley.Fields("Descripcio")
        objExcel.ActiveSheet.Cells(v, h + 2) = "" & mytabley.Fields("Unidad")
        objExcel.ActiveSheet.Cells(v, h + 3) = "" & mytabley.Fields("Factor")
        objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytabley.Fields("Cantidad")
        objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytabley.Fields("Precio")
        objExcel.ActiveSheet.Cells(v, h + 6) = "" & mytabley.Fields("Total")
        v = v + 1
        mytabley.MoveNext
    Loop
    objExcel.ActiveSheet.Cells(v, h + 6) = Format(Val(total), "0.00")
    mytabley.Close
    mytablex.Close
     
    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    Exit Sub
cmd5612_err:
    MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub cmdAddEntry_Click()

    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub

    inicializa_insumo
    Frame2.Visible = True
    Frame2.Caption = "NUEVO"
    productoi.Enabled = True
    cmdBuscarProducto.Enabled = True
    productoi.SetFocus

    cmdBuscarProducto_Click

End Sub

Sub inicializa_insumo()
    lineai = ""
    xtotal = ""
    productoi = ""
    descripcioi = ""
    unidad = ""
    factor = ""
    cantidad = ""
    precio = ""
    'linea = ""
    t1 = ""
    t2 = ""
    t3 = ""
    t4 = ""
    t5 = ""
    t6 = ""
    t7 = ""
    t8 = ""
    t9 = ""
    t10 = ""
    t11 = ""
    t12 = ""
    t13 = ""
    t14 = ""
    t15 = ""
    t16 = ""
    nt1 = ""
    nt2 = ""
    nt3 = ""
    nt4 = ""
    nt5 = ""
    nt6 = ""
    nt7 = ""
    nt8 = ""
    nt9 = ""
    nt10 = ""
    nt11 = ""
    nt12 = ""
    nt13 = ""
    nt14 = ""
    nt15 = ""
    nt16 = ""

End Sub

Private Sub cmdBuscarProducto_Click()
    consulta_producto

End Sub

Private Sub cmdSave_Click()

    Dim found As Integer

    If Len(Trim("" & productoi)) = 0 Then
        productoi.SetFocus
        Exit Sub

    End If

    If Val("" & cantidad) = 0 Then
        cantidad.SetFocus
        Exit Sub

    End If

    If Len(Trim("" & lineai)) > 0 Then
        suma_linea

    End If

    ''' 11/12/2017 SubReceta
    'xtotal = Format(Val(cantidad) * Val(precio), "0.00")
    xtotal = Format(Val(cantidad) * Val(precio), "0.0000")
    ''' 11/12/2017 SubReceta

    found = graba_receta()

    '13/08/2018 Integración FE - Pizzeria
    ''' 11/12/2017 SubReceta
    'Actualiza_CostosReceta
    ''' 11/12/2017 SubReceta
    '13/08/2018 Integración FE - Pizzeria

    If found = 0 Then
        Exit Sub

    End If

    sql_receta
    lfdo33_Click

End Sub

'13/08/2018 Integración FE - Pizzeria
'' 11/12/2017 SubReceta
Sub grabandoSubReceta(mytablex As ADODB.Recordset)

    Dim mytable11 As New ADODB.Recordset

    If mytable11.State = 1 Then mytable11.Close
    mytable11.Open "SELECT * FROM receta where producto='" & productoi & "'", cn, adOpenDynamic, adLockOptimistic

    If mytable11.RecordCount > 0 Then
        Do

            If mytable11.EOF Then Exit Do
            mytablex.AddNew

            mytablex.Fields("nro") = mytable11.Fields("nro")
            mytablex.Fields("producto") = Trim(producto)
            mytablex.Fields("productoi") = mytable11.Fields("productoi")
            mytablex.Fields("descripcio") = mytable11.Fields("descripcio")
            mytablex.Fields("linea") = mytable11.Fields("producto")
            mytablex.Fields("unidad") = mytable11.Fields("unidad")
            mytablex.Fields("factor") = mytable11.Fields("factor")
            mytablex.Fields("precio") = mytable11.Fields("precio")
            mytablex.Fields("t1") = mytable11.Fields("t1")
            mytablex.Fields("t2") = mytable11.Fields("t2")
            mytablex.Fields("t3") = mytable11.Fields("t3")
            mytablex.Fields("t4") = mytable11.Fields("t4")
            mytablex.Fields("t5") = mytable11.Fields("t5")
            mytablex.Fields("t6") = mytable11.Fields("t6")
            mytablex.Fields("t7") = mytable11.Fields("t7")
            mytablex.Fields("t8") = mytable11.Fields("t8")
            mytablex.Fields("t9") = mytable11.Fields("t9")
            mytablex.Fields("t10") = mytable11.Fields("t10")
            mytablex.Fields("t11") = mytable11.Fields("t11")
            mytablex.Fields("t12") = mytable11.Fields("t12")
            mytablex.Fields("t13") = mytable11.Fields("t13")
            mytablex.Fields("t14") = mytable11.Fields("t14")
            mytablex.Fields("t15") = mytable11.Fields("t15")
            mytablex.Fields("t16") = mytable11.Fields("t16")

            'Cambios Pizzeria 24/05/2018
            '          Call obtieneTipoReceta(tipodereceta)
            '          If tipodereceta = "P" Then ' Produccion
            '                mytablex.Fields("cantidad") = Val("" & cantidad)
            '                mytablex.Fields("total") = mytable11.Fields("total")
            '          Else
            mytablex.Fields("cantidad") = Format(mytable11.Fields("cantidad") * Val("" & cantidad), "0.00000")
            mytablex.Fields("total") = Format(mytablex.Fields("cantidad") * mytable11.Fields("precio"), "0.00000")
            ' End If
            'Cambios Pizzeria 24/05/2018
          
            mytablex.Update
            mytable11.MoveNext
        Loop

    End If

    mytable11.Close

End Sub

'13/08/2018 Integración FE - Pizzeria

'13/08/2018 Integración FE - Pizzeria
Sub ActualizaSubRecetaProduccion()

    Dim mytable11 As New ADODB.Recordset

    If mytable11.State = 1 Then mytable11.Close
    mytable11.Open "SELECT * FROM receta  where linea='" & producto & "' and productoi='" & productoi & "'", cn, adOpenDynamic, adLockOptimistic

    If mytable11.RecordCount > 0 Then
           
        Do

            If mytable11.EOF Then Exit Do
            mytable11.Fields("unidad") = Trim(unidad)
            mytable11.Fields("factor") = Val(factor)
            mytable11.Fields("precio") = Val(precio)
            
            mytable11.Fields("cantidad") = Val(cantidad)
            mytable11.Fields("total") = Val("" & precio) * Val("" & cantidad)
            mytable11.Update
            mytable11.MoveNext
        Loop

    End If

    mytable11.Close
   
    Dim mytablex As New ADODB.Recordset

    Dim suma     As Double

    suma = 0
   
    Set DBGrid2.DataSource = mytablev
    Do

        If mytablev.EOF Then Exit Do
        suma = suma + Val("" & mytablev.Fields("cantidad")) * Val("" & mytablev.Fields("precio"))
        mytablev.MoveNext
    Loop
    total = Format(suma, "0.00")

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM receta  where  productoI='" & producto & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
   
        Do

            If mytablex.EOF Then Exit Do
            mytablex.Fields("precio") = Val(total)
            mytablex.Fields("total") = Val("" & total) * mytablex.Fields("cantidad")
            mytablex.Update
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

End Sub

Sub ActualizaSubRecetaEstandar()

    Dim mytable11       As New ADODB.Recordset

    Dim mytablecan      As New ADODB.Recordset

    Dim cantidadInicial As String
       
    If mytablecan.State = 1 Then mytablecan.Close
    mytablecan.Open "SELECT * FROM receta  where productoi='" & producto & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablecan.RecordCount > 0 Then
        cantidadInicial = Val(mytablecan.Fields("cantidad"))

    End If

    mytablecan.Close
       
    If mytable11.State = 1 Then mytable11.Close
    mytable11.Open "SELECT * FROM receta  where linea='" & producto & "' and productoi='" & productoi & "'", cn, adOpenDynamic, adLockOptimistic

    If mytable11.RecordCount > 0 Then
           
        Do

            If mytable11.EOF Then Exit Do
            mytable11.Fields("unidad") = Trim(unidad)
            mytable11.Fields("factor") = Val(factor)
            mytable11.Fields("precio") = Val(precio)
            
            mytable11.Fields("cantidad") = cantidadInicial * Val(cantidad)
            mytable11.Fields("total") = Val("" & precio) * Val(mytable11.Fields("cantidad"))
            mytable11.Update
            mytable11.MoveNext
        Loop

    End If

    mytable11.Close
   
    Dim mytablex As New ADODB.Recordset

    Dim suma     As Double

    suma = 0
   
    Set DBGrid2.DataSource = mytablev
    Do

        If mytablev.EOF Then Exit Do
        suma = suma + Val("" & mytablev.Fields("cantidad")) * Val("" & mytablev.Fields("precio"))
        mytablev.MoveNext
    Loop
    total = Format(suma, "0.00")

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM receta  where  productoI='" & producto & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
   
        Do

            If mytablex.EOF Then Exit Do
            mytablex.Fields("precio") = Val(total)
            mytablex.Fields("total") = Val("" & total) * mytablex.Fields("cantidad")
            mytablex.Update
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

End Sub

Function obtieneTipoReceta(ByRef tipo As String)

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
 
    mytablex.Open "SELECT tiporeceta FROM parame where codigo =1 ", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        If IsNull(mytablex.Fields("tiporeceta")) Then
            tipo = ""
        Else
            tipo = mytablex.Fields("tiporeceta")

        End If

    End If

    mytablex.Close

End Function

Sub Actualiza_CostosReceta()

    On Error GoTo cmd9093_err

    cn.Execute ("update receta set precio=" & Val(precio) & ",total='" & Val(total) & "' where productoi='" & productoi & "'")
    Exit Sub
cmd9093_err:
    MsgBox "Aviso en actualiza receta " + error$, 48, "Aviso"
    Exit Sub

End Sub

'' 11/12/2017 SubReceta
'13/08/2018 Integración FE - Pizzeria

Private Sub Command1_Click()
    ejecuta 1

End Sub

Private Sub Command1X_Click()
    lfdo33_Click

End Sub

Private Sub Command2_Click()
    lfdo33_Click

End Sub

Private Sub Command2X_Click()
    inicializa_insumo
    productoi.SetFocus

End Sub

Private Sub Command3_Click()
    lfdo33_Click

End Sub

Private Sub copia1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0

End Sub

Private Sub copia2_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0

End Sub

'13/08/2018 Integración FE - Pizzeria
Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim found    As Integer

    Dim rsexiste As New ADODB.Recordset

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "2" Then
            'xproducto = dbGrid1.columns(1)
            'Frame1.Visible = False
            'xproducto.SetFocus
            Exit Sub

        End If

        'AQUI YA
        If opcion1 = "1" Then

            'Cambios Pizzeria 24/05/2018
            '   rsexiste.Open "SELECT * FROM " & tiporeceta & " where  nro='" & Trim(nro) & "' and producto='" & Trim(producto) & "' and productoi='" & Trim(dbGrid1.columns(1)) & "'", cn, adOpenKeyset, adLockOptimistic
            '
            '   If rsexiste.RecordCount > 0 Then  'si existe
            '      MsgBox "Ya existe insumo ", 48, "Aviso"
            '      Exit Sub
            '   End If
            'Cambios Pizzeria 24/05/2018

            found = busca_producto("" & dbGrid1.columns(1))
            cantidad = "1"
   
            ''' 11/12/2017 SubReceta
   
            If platos > 1 Then
                platos2.Visible = True
                platos2.Caption = "(Para " & Val(platos) & " Und.) "
     
                cantidad = Format(Val(cantidad) / Val(platos), "0.0000000")
            Else
                platos2.Visible = False

            End If
  
            ''' 11/12/2017 SubReceta

            'lineai = dbGrid1.columns(5)
            'found = busca_linea("" & lineai)
            Frame1.Visible = False
            Frame1.Enabled = False
            'xtotal = Format(Val(cantidad) * Val(precio), "0.00")
            cantidad.SetFocus

            If Len(lineai) > 0 Then
                t1.SetFocus

            End If

        End If
   
        ''' 11/12/2017 SubReceta
        'xtotal = Format(Val(cantidad) * Val(precio), "0.00")
        xtotal = Format(Val(cantidad) * Val(precio), "0.0000")

        ''' 11/12/2017 SubReceta
    End If

End Sub

'13/08/2018 Integración FE - Pizzeria

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
        ejecuta 0
         
    End If

End Sub

Private Sub DBGrid2_DblClick()

    Dim found As Integer

    On Error GoTo cmd23_err

    inicializa_insumo
    'If Len("" & mytablev.fields("")) = 0 Then Exit Sub
    productoi = "" & mytablev.Fields("productoi")
    descripcioi = "" & mytablev.Fields("descripcio")
    lineai = "" & mytablev.Fields("linea")
    precio = "" & mytablev.Fields("precio")
    unidad = "" & mytablev.Fields("unidad")
    factor = "" & mytablev.Fields("factor")
    cantidad = "" & mytablev.Fields("cantidad")
    t1 = "" & "" & mytablev.Fields("t1")
    t2 = "" & "" & mytablev.Fields("t2")
    t3 = "" & "" & mytablev.Fields("t3")
    t4 = "" & "" & mytablev.Fields("t4")
    t5 = "" & "" & mytablev.Fields("t5")
    t6 = "" & "" & mytablev.Fields("t6")
    t7 = "" & "" & mytablev.Fields("t7")
    t8 = "" & "" & mytablev.Fields("t8")
    t9 = "" & "" & mytablev.Fields("t9")
    t10 = "" & "" & mytablev.Fields("t10")
    t11 = "" & "" & mytablev.Fields("t11")
    t12 = "" & "" & mytablev.Fields("t12")
    t13 = "" & "" & mytablev.Fields("t13")
    t14 = "" & "" & mytablev.Fields("t14")
    t15 = "" & "" & mytablev.Fields("t15")
    t16 = "" & "" & mytablev.Fields("t16")
    xtotal = "" & "" & mytablev.Fields("total")
    Frame2.Caption = "MODIFICA"
    found = busca_linea("" & lineai)
    Frame2.Visible = True
    productoi.Enabled = False
    cmdBuscarProducto.Enabled = False

    cantidad.SetFocus
    Exit Sub
cmd23_err:
    Exit Sub

End Sub

Private Sub DBGrid2_KeyUp(KeyCode As Integer, Shift As Integer)

    On Error GoTo cmd46_err

    If KeyCode = &H2E Then  'borrar linea

        'Data2.Recordset.Delete
        'Data2.Refresh
    End If

    Exit Sub
cmd46_err:
    Exit Sub

End Sub

Private Sub dki232_Click()

    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub

    cmdAddEntry_Click

End Sub

Private Sub dmi33_Click()

    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub

    DBGrid2_DblClick

End Sub

Private Sub fd2b34_Click()

    On Error GoTo cmd56_err

    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub

    If MsgBox("Desea Borrar " & DBGrid2.columns(0), 1, "Aviso") <> 1 Then Exit Sub
    cn.Execute ("DELETE   FROM " & tiporeceta & " WHERE nro='" & Trim(nro) & "' and producto='" & producto & "' and productoi='" & DBGrid2.columns(0) & "'")

    '' 11/12/2017 SubReceta
    cn.Execute ("DELETE   FROM " & tiporeceta & " WHERE nro='" & Trim(nro) & "' and producto='" & producto & "' and linea='" & DBGrid2.columns(0) & "'")
    '' 11/12/2017 SubReceta

    sql_receta
    Exit Sub
cmd56_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"

End Sub

Private Sub Form_Activate()
    Frame1.Top = 10: Frame1.Left = 10
    Frame2.Top = 10: Frame2.Left = 10
    nro_KeyPress 13
    cargas_iniciales

End Sub

Sub cargas_iniciales()

    Dim cad      As String

    Dim mytablex As New ADODB.Recordset

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

End Sub

Function extra_loquesea1(buf As String) As String

    Dim j

    Dim buf1 As String

    buf1 = ""

    If InStr(buf, "|") > 0 Then
        j = InStr(buf, "|")
        buf1 = Mid$(buf, j + 1, Len(buf) - (j))
    Else
        buf1 = buf

    End If

    extra_loquesea1 = buf1

End Function

Sub limpia_linea()
    lineai = ""
    nlinea = ""
    nt1 = ""
    nt2 = ""
    nt3 = ""
    nt4 = ""
    nt5 = ""
    nt6 = ""
    nt7 = ""
    nt8 = ""
    nt9 = ""
    nt10 = ""
    nt11 = ""
    nt12 = ""
    nt13 = ""
    nt14 = ""
    nt15 = ""
    nt16 = ""

End Sub

Function busca_linea(buf As String)

    Dim mytablex As New ADODB.Recordset

    limpia_linea

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from linea where linea='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF = True And mytablex.BOF = True Then
        mytablex.Close
        Exit Function

    End If

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
    mytablex.Close
 
End Function

Private Sub lfdo33_Click()

    If opcion1 = "1" Then
        If Frame1.Visible = True Then
            Frame1.Visible = False
            productoi.SetFocus
            Exit Sub

        End If

    End If

    If opcion1 = "2" Then
        If Frame1.Visible = True Then
            Frame1.Visible = False
            'xproducto.SetFocus
            Exit Sub

        End If

    End If

    'If Frame3.Visible = True Then
    '   Frame3.Visible = False
    '   Exit Sub
    'End If

    If Frame2.Visible = True Then
        Frame2.Visible = False
        DBGrid2.SetFocus
        Exit Sub

    End If

    treceta.Hide
    Unload treceta

End Sub

Private Sub nro_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(nro) = 0 Then
        MsgBox "Debe existir al menos un valor ", 48, "Aviso"
        nro.SetFocus
        Exit Sub

    End If

    sql_receta

End Sub

'13/08/2018 Integración FE - Pizzeria
Sub sql_receta()

    Dim suma As Double

    suma = 0

    If mytablev.State = 1 Then mytablev.Close
    '' 11/12/2017 SubReceta
    'mytablev.Open "select * from " & tiporeceta & " where nro='" & nro & "' and producto='" & producto & "'", cn, adOpenStatic, adLockOptimistic
    mytablev.Open "select * from " & tiporeceta & " where  LINEA='' AND nro='" & nro & "' and producto='" & producto & "'", cn, adOpenStatic, adLockOptimistic
    '' 11/12/2017 SubReceta
   
    Set DBGrid2.DataSource = mytablev
    Do

        If mytablev.EOF Then Exit Do
        suma = suma + Val("" & mytablev.Fields("cantidad")) * Val("" & mytablev.Fields("precio"))
        mytablev.MoveNext
    Loop
      
    '' 11/12/2017 SubReceta
    'total = Format(suma, "0.00")
    'graba_costoproduccion Trim("" & producto)
    total = Format(suma, "0.00000000")

    If total > 0 Then

        '''' 11/12/2017 SubReceta
        'graba_costoproduccion Trim("" & producto)
        If OpcionActualizaCostoReceta("" & producto) = "S" Then
            graba_costoproduccion Trim("" & producto)

        End If

        '''' 11/12/2017 SubReceta
    End If

    '' 11/12/2017 SubReceta
    DBGrid2.SetFocus

End Sub

'13/08/2018 Integración FE - Pizzeria

Private Sub precio_KeyPress(KeyAscii As Integer)

    '' 11/12/2017 SubReceta
    'xtotal = Format(Val(cantidad) * Val(precio), "0.00")
    xtotal = Format(Val(cantidad) * Val(precio), "0.0000")

    '' 11/12/2017 SubReceta
End Sub

'13/08/2018 Integración FE - Pizzeria
Private Sub productoi_KeyPress(KeyAscii As Integer)

    Dim rsexiste As New ADODB.Recordset

    Dim found    As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(productoi) = 0 Then
        inicializa_insumo
        Exit Sub

    End If

    ' AQUI YA
    If Frame2 = "NUEVO" Then

        'Cambios Pizzeria 24/05/2018
        'rsexiste.Open "SELECT * FROM " & tiporeceta & " where  nro='" & Trim(nro) & "' and producto='" & Trim(producto) & "' and productoi='" & Trim(productoi) & "'", cn, adOpenKeyset, adLockOptimistic
        'If rsexiste.RecordCount > 0 Then  'si existe
        '   inicializa_insumo
        '   MsgBox "Ya existe insumo ", 48, "Aviso"
        '   Exit Sub
        'End If
        'Cambios Pizzeria 24/05/2018
    End If

    found = busca_producto("" & productoi)

    If found = 0 Then
        MsgBox "No existe Producto", 48, "Aviso"
        inicializa_insumo
      
        '' 11/12/2017 SubReceta
        'xtotal = Format(Val(cantidad) * Val(precio), "0.00")
        xtotal = Format(Val(cantidad) * Val(precio), "0.0000")
        '' 11/12/2017 SubReceta
      
        Exit Sub

    End If

    cantidad.SetFocus

    If Len(lineai) > 0 Then
        t1.SetFocus

    End If
   
    '' 11/12/2017 SubReceta
    'xtotal = Format(Val(cantidad) * Val(precio), "0.00")
    xtotal = Format(Val(cantidad) * Val(precio), "0.00")

    '' 11/12/2017 SubReceta
End Sub

'13/08/2018 Integración FE - Pizzeria

Private Sub productoi_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_producto

    End If

End Sub

Sub consulta_producto()
    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "Producto"

    ''''12/08/2017 kenyo Busqueda por codigo de barras en recetas
    Combo1.AddItem "Barras"
    ''''12/08/2017 kenyo Busqueda por codigo de barras en recetas

    Combo1.ListIndex = 0
    Frame1.Enabled = True
    Frame1.Visible = True
    buffer = ""
    opcion1 = "1"
    buffer.SetFocus
    Command1_Click

End Sub

Private Sub t1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    suma_linea
    t2.SetFocus

End Sub

Private Sub t10_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    suma_linea
    t11.SetFocus

End Sub

Private Sub t11_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    suma_linea
    t12.SetFocus

End Sub

Private Sub t12_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    suma_linea
    t13.SetFocus

End Sub

Private Sub t13_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    suma_linea
    t14.SetFocus

End Sub

Private Sub t14_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    suma_linea
    t15.SetFocus

End Sub

Private Sub t15_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    suma_linea
    t16.SetFocus

End Sub

Private Sub t16_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    suma_linea

End Sub

Private Sub t2_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    suma_linea
    t3.SetFocus

End Sub

Private Sub t3_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    suma_linea
    t4.SetFocus

End Sub

Private Sub t4_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    suma_linea
    t5.SetFocus

End Sub

Private Sub t5_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    suma_linea
    t6.SetFocus

End Sub

Private Sub t6_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    suma_linea
    t7.SetFocus

End Sub

Private Sub t7_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    suma_linea
    t8.SetFocus

End Sub

Private Sub t8_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    suma_linea
    t9.SetFocus

End Sub

Private Sub t9_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    suma_linea
    t10.SetFocus

End Sub

'13/08/2018 Integración FE - Pizzeria
Function graba_receta()

    Dim rsexiste As New ADODB.Recordset

    'MsgBox Frame2
    If Frame2 = "NUEVO" Then
        'Cambios Pizzeria 24/05/2018
        ' rsexiste.Open "SELECT * FROM " & tiporeceta & " where  nro='" & Trim(nro) & "' and producto='" & Trim(producto) & "' and productoi='" & Trim(productoi) & "'", cn, adOpenKeyset, adLockOptimistic
        ' If rsexiste.RecordCount > 0 Then  'si existe
        '     MsgBox "Producto ya existe,", 48, "Aviso"
        ' Exit Function
        ' End If
        ' rsexiste.Close
        'Cambios Pizzeria 24/05/2018
      
        mytablev.AddNew
        grabando mytablev
      
        mytablev.Update
        graba_receta = 1
      
        '' 11/12/2017 SubReceta
        grabandoSubReceta mytablev
        '' 11/12/2017 SubReceta

        Exit Function

    End If

    If Frame2 = "MODIFICA" Then
        grabando mytablev
        mytablev.Update
        graba_receta = 1
      
        '' 11/12/2017 SubReceta
        '      Call obtieneTipoReceta(tipodereceta)
        '      If tipodereceta = "P" Then ' Produccion
        '            ActualizaSubRecetaProduccion
        '      Else
        '            ActualizaSubRecetaEstandar
        '      End If
        '' 11/12/2017 SubReceta
      
    End If

End Function

Sub graba_costoproduccion(buf As String)

    Dim mytablex As New ADODB.Recordset

    If Val(platos) <= 0 Then
        platos = "1"

    End If

    mytablex.Open "select * from producto where producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
      
        ''' 11/12/2017 SubReceta
        'mytablex.Fields("costou") = Val(total) / Val(platos)
        mytablex.Fields("costou") = Val(total)
        ''' 11/12/2017 SubReceta
      
        mytablex.Update

    End If

    mytablex.Close

End Sub

'13/08/2018 Integración FE - Pizzeria

Sub grabando(mytablex As ADODB.Recordset)

    Dim cad As String

    mytablex.Fields("nro") = Trim(nro)
    mytablex.Fields("producto") = Trim(producto)
    mytablex.Fields("productoi") = Trim(productoi)
    mytablex.Fields("descripcio") = Trim(descripcioi)
    mytablex.Fields("linea") = Trim(lineai)
    mytablex.Fields("unidad") = Trim(unidad)
    mytablex.Fields("factor") = Val(factor)
    mytablex.Fields("precio") = Val(precio)
    mytablex.Fields("t1") = Val(t1)
    mytablex.Fields("t2") = Val(t2)
    mytablex.Fields("t3") = Val(t3)
    mytablex.Fields("t4") = Val(t4)
    mytablex.Fields("t5") = Val(t5)
    mytablex.Fields("t6") = Val(t6)
    mytablex.Fields("t7") = Val(t7)
    mytablex.Fields("t8") = Val(t8)
    mytablex.Fields("t9") = Val(t9)
    mytablex.Fields("t10") = Val(t10)
    mytablex.Fields("t11") = Val(t11)
    mytablex.Fields("t12") = Val(t12)
    mytablex.Fields("t13") = Val(t13)
    mytablex.Fields("t14") = Val(t14)
    mytablex.Fields("t15") = Val(t15)
    mytablex.Fields("t16") = Val(t16)
    mytablex.Fields("cantidad") = Val(cantidad)
    mytablex.Fields("total") = Val("" & precio) * Val("" & cantidad)

End Sub

Sub suma_linea()

    Dim sdx As Double

    sdx = Val(t1) + Val(t2) + Val(t3) + Val(t4) + Val(t5) + Val(t6) + Val(t7) + Val(t8) + Val(t9) + Val(t10) + Val(t11) + Val(t12) + Val(t13) + Val(t14) + Val(t15) + Val(t16)
    cantidad = Format(sdx, "0.0000")
   
    '' 11/12/2017 SubReceta
    'xtotal = Format(Val(cantidad) * Val(precio), "0.00")
    xtotal = Format(Val(cantidad) * Val(precio), "0.0000")

    '' 11/12/2017 SubReceta
End Sub

'13/08/2018 Integración FE - Pizzeria
Function busca_producto(buf As String)

    Dim found    As Integer

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    productoi = ""
    descripcioi = ""
    unidad = ""
    factor = ""
    precio = ""
    lineai = ""
    nlinea = ""
    cantidad = "1"
    mytablex.Open "SELECT * FROM producto where  producto='" & Trim(buf) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
      
        OpcionTipoCostoReceta
      
        productoi = "" & mytablex.Fields("producto")
        descripcioi = "" & mytablex.Fields("descripcio")
        unidad = "" & mytablex.Fields("unidadP")
      
        ''' 11/12/2017 SubReceta
        platosi = "" & mytablex.Fields("platos")
        ''' 11/12/2017 SubReceta
      
        factor = "1"

        If Val("" & mytablex.Fields("factorP")) > 0 Then
            factor = Format(1 / Val("" & mytablex.Fields("factorP")), "0.00000")

        End If
      
        '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
        'sdx = Val("" & mytablex.Fields("costou")) * Val(factor)
        If OpcionTipoCostoReceta() = "CP" Then
            sdx = Val("" & mytablex.Fields("costop")) * Val(factor)
        Else
            sdx = Val("" & mytablex.Fields("costou")) * Val(factor)

        End If

        '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
   
        '' 11/12/2017 SubReceta
        'precio = Format(sdx, "0.00")
        precio = Format(sdx, "0.00000")
        '' 11/12/2017 SubReceta
      
        '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
        'lineai = "" & mytablex.Fields("costou")
        If OpcionTipoCostoReceta() = "CP" Then
            lineai = "" & mytablex.Fields("costop")
        Else
            lineai = "" & mytablex.Fields("costou")

        End If

        '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
        found = busca_linea("" & lineai)
        busca_producto = 1

    End If

    mytablex.Close

End Function

'13/08/2018 Integración FE - Pizzeria

Sub ir_inicio()

End Sub

Private Sub xproducto_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub xproducto_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_producto1

    End If

End Sub

Sub consulta_producto1()
    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "Producto"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    buffer = ""
    opcion1 = "2"
    buffer.SetFocus
    Command1_Click

End Sub

'13/08/2018 Integración FE - Pizzeria
''' 11/12/2017 SubReceta
'11/06/2018 Actualiza Precio Promedio Ponderado Masivo
'Sub ejecuta(sw As Integer)
'Dim rconsulta As New ADODB.Recordset
'Dim cad As String
'If opcion1 = "1" Then  'bodega
'
'   If Len(buffer) = 0 Then
'
'      ''' 11/12/2017 SubReceta
'            'cad = "select Descripcio,Producto,Unidadp,Factorp,Costou,Linea,familia from producto "
'      cad = "select Descripcio,Producto,Unidadp,Factorp,Costou,Linea,familia,platos from producto "
'       ''' 11/12/2017 SubReceta
'
'
'         If familia <> "%" Then
'         cad = cad & " where producto.familia like '" & extra_loquesea1(familia) & "%'"
'          End If
'   End If
'
'
'   If Len(buffer) > 0 Then
'
'    ''' 11/12/2017 SubReceta
'    'cad = "select Descripcio,Producto,Unidadp,Factorp,Costou,Linea,familia from producto where " & Combo1 & " like '" & buffer & "%'"
'      cad = "select Descripcio,Producto,Unidadp,Factorp,Costou,Linea,familia,platos from producto where " & Combo1 & " like '%" & buffer & "%'"
'    ''' 11/12/2017 SubReceta
'
'        'kenyo receta busqueda por receta
'        If familia <> "%" Then
'            cad = cad & " and producto.familia like '" & extra_loquesea1(familia) & "'"
'        End If
'
'
'   End If
'   If rconsulta.State = 1 Then rconsulta.Close
'   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
'   If rconsulta.EOF = True And rconsulta.BOF = True Then
'      rconsulta.Close
'      buffer.SetFocus
'      Exit Sub
'   End If
'   Set dbGrid1.DataSource = rconsulta
'               dbGrid1.columns(0).Width = 4000
'               dbGrid1.columns(1).Width = 2000
'               dbGrid1.columns(2).Width = 1000
'               dbGrid1.columns(3).Width = 1000
'               dbGrid1.columns(4).Width = 500
'               dbGrid1.columns(5).Width = 1000
'   If sw = 1 Then
'      dbGrid1.SetFocus
'
'   End If
'   Exit Sub
'End If
'End Sub
Sub ejecuta(sw As Integer)

    Dim rconsulta As New ADODB.Recordset

    Dim cad       As String

    If opcion1 = "1" Then  'bodega
   
        If Len(buffer) = 0 Then
   
            cad = "select Descripcio,Producto,Unidadp,Factorp,Costou, Costop, Linea,familia,platos from producto "

            If familia <> "%" Then
                cad = cad & " where producto.familia like '" & extra_loquesea1(familia) & "%'"

            End If

        End If
   
        If Len(buffer) > 0 Then
   
            cad = "select Descripcio,Producto,Unidadp,Factorp,Costou,Costop,Linea,familia,platos from producto where " & Combo1 & " like '%" & buffer & "%'"

            If familia <> "%" Then
                cad = cad & " and producto.familia like '" & extra_loquesea1(familia) & "'"

            End If
   
        End If

        If rconsulta.State = 1 Then rconsulta.Close
        rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

        If rconsulta.EOF = True And rconsulta.BOF = True Then
            rconsulta.Close
            buffer.SetFocus
            Exit Sub

        End If

        Set dbGrid1.DataSource = rconsulta
        dbGrid1.columns(0).Width = 4150
        dbGrid1.columns(1).Width = 2000
        dbGrid1.columns(2).Width = 800
        dbGrid1.columns(3).Width = 800
        dbGrid1.columns(4).Width = 900
        dbGrid1.columns(5).Width = 900
        dbGrid1.columns(6).Width = 0
        dbGrid1.columns(7).Width = 1000
        dbGrid1.columns(8).Width = 0
               
        If sw = 1 Then
            dbGrid1.SetFocus
      
        End If

        Exit Sub

    End If

End Sub

'11/06/2018 Actualiza Precio Promedio Ponderado Masivo
''' 11/12/2017 SubReceta
'13/08/2018 Integración FE - Pizzeria

