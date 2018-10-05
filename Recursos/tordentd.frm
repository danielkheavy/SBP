VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tordentd 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Detalle Orden Trabajo"
   ClientHeight    =   10080
   ClientLeft      =   165
   ClientTop       =   -615
   ClientWidth     =   15810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   15810
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000D&
      Caption         =   "Formula"
      Height          =   8415
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   12735
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         Height          =   735
         Left            =   11640
         TabIndex        =   39
         Top             =   240
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid dbgrid3 
         Height          =   7215
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   12726
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
      Begin VB.Label nround 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   8760
         TabIndex        =   40
         Top             =   7680
         Width           =   2775
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consulta"
      Height          =   9975
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   15015
      Begin VB.TextBox Text1 
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
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
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
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Filtrar"
         Height          =   495
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   2175
      End
      Begin MSDataGridLib.DataGrid dbgrid13 
         Height          =   8895
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   15690
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
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Frame2"
      Height          =   8895
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   14295
      Begin VB.TextBox unidadf 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
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
         Left            =   6600
         MaxLength       =   6
         TabIndex        =   51
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox factorf 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
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
         Left            =   5160
         MaxLength       =   6
         TabIndex        =   49
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Formula Necesaria"
         Height          =   4815
         Left            =   120
         TabIndex        =   44
         Top             =   3720
         Visible         =   0   'False
         Width           =   14055
         Begin VB.CommandButton Command2 
            Caption         =   "CargaFormula"
            Height          =   615
            Left            =   240
            TabIndex        =   47
            Top             =   3960
            Width           =   1335
         End
         Begin VB.CommandButton Command5 
            Caption         =   "CargaStock"
            Height          =   615
            Left            =   1680
            TabIndex        =   46
            Top             =   3960
            Width           =   1335
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Refresca"
            Height          =   615
            Left            =   3120
            TabIndex        =   45
            Top             =   3960
            Width           =   1335
         End
         Begin MSDataGridLib.DataGrid dbgrid5 
            Height          =   3615
            Left            =   240
            TabIndex        =   48
            Top             =   240
            Width           =   13695
            _ExtentX        =   24156
            _ExtentY        =   6376
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   2
            RowHeight       =   29
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
            ColumnCount     =   12
            BeginProperty Column00 
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
            BeginProperty Column01 
               DataField       =   "Descripcio"
               Caption         =   "Descripcio"
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
               DataField       =   "Explosion"
               Caption         =   "Ex"
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
               DataField       =   "Formula"
               Caption         =   "Formula"
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
               DataField       =   "UNidad"
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
            BeginProperty Column05 
               DataField       =   "factor"
               Caption         =   "Factor"
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
               DataField       =   "Cantidad"
               Caption         =   "Cantidad"
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
               DataField       =   "Stock"
               Caption         =   "Stock"
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
               DataField       =   "Faltante"
               Caption         =   "Faltante"
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
               DataField       =   "Costo"
               Caption         =   "Costo"
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
               DataField       =   "TotalCosto"
               Caption         =   "TotalCosto"
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
               DataField       =   "OrdenTrabajod"
               Caption         =   "OrdenTrabajod"
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
                  ColumnWidth     =   1184.882
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   5295.118
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   615.118
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   750.047
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   794.835
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   750.047
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   870.236
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   989.858
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   959.811
               EndProperty
               BeginProperty Column10 
               EndProperty
               BeginProperty Column11 
               EndProperty
            EndProperty
         End
         Begin VB.Label total 
            BorderStyle     =   1  'Fixed Single
            Height          =   495
            Left            =   11520
            TabIndex        =   53
            Top             =   4080
            Width           =   1815
         End
      End
      Begin VB.TextBox formula 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
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
         TabIndex        =   33
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox bodega 
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
         TabIndex        =   24
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   10560
         Picture         =   "tordentd.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   480
         Width           =   1470
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   10560
         Picture         =   "tordentd.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Imprimir todo"
         Top             =   1560
         Width           =   1470
      End
      Begin VB.TextBox descripcio 
         Enabled         =   0   'False
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
         MaxLength       =   100
         TabIndex        =   16
         Top             =   600
         Width           =   6015
      End
      Begin VB.TextBox producto 
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
         TabIndex        =   15
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox unidad 
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
         TabIndex        =   14
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox factor 
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
         MaxLength       =   4
         TabIndex        =   13
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox cantidad 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   12
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lo que se va a Producir"
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
         TabIndex        =   52
         Top             =   1320
         Width           =   8175
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Formula Para"
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
         Left            =   3720
         TabIndex        =   50
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Refresca"
         Height          =   375
         Left            =   3720
         TabIndex        =   43
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Stock"
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
         TabIndex        =   42
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label stock 
         BackColor       =   &H00FFFF00&
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
         Height          =   375
         Left            =   2280
         TabIndex        =   41
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3720
         Picture         =   "tordentd.frx":1194
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   375
      End
      Begin VB.Image Image4 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4200
         Picture         =   "tordentd.frx":149E
         Stretch         =   -1  'True
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro Formula"
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
         TabIndex        =   32
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Almacen"
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
         TabIndex        =   25
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
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
         TabIndex        =   23
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto"
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
         TabIndex        =   22
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unidad"
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
         TabIndex        =   21
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Factor"
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
         TabIndex        =   20
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CantidadProducir"
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
         TabIndex        =   19
         Top             =   3240
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   12435
      TabIndex        =   2
      Top             =   0
      Width           =   12495
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
         Picture         =   "tordentd.frx":17A8
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Grabar registro"
         Top             =   0
         Width           =   735
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
         Left            =   7560
         MaxLength       =   10
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   120
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "&Buscar"
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
         Left            =   10560
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00E0E0E0&
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
         Picture         =   "tordentd.frx":29BA
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   2760
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tordentd.frx":3BCC
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "tordentd.frx":4DDE
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "tordentd.frx":5FF0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   12495
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7335
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   12938
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
         ColumnCount     =   10
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "Producto"
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
         BeginProperty Column02 
            DataField       =   "Formula"
            Caption         =   "Formula"
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
            DataField       =   "Unidad"
            Caption         =   "Und"
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
         BeginProperty Column04 
            DataField       =   "Factor"
            Caption         =   "Factor"
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
         BeginProperty Column05 
            DataField       =   "Cantidad"
            Caption         =   "Cant"
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
         BeginProperty Column06 
            DataField       =   "Bodega"
            Caption         =   "Almacen"
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
         BeginProperty Column07 
            DataField       =   "Ordentrabajo"
            Caption         =   "Ordentrabajo"
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
         BeginProperty Column08 
            DataField       =   "Formula"
            Caption         =   "Formula"
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
         BeginProperty Column09 
            DataField       =   "OrdenTrabajod"
            Caption         =   "Ordentrabajod"
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
               ColumnWidth     =   6029.858
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
            EndProperty
         EndProperty
      End
      Begin VB.Label ornumero 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   4080
         TabIndex        =   36
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label orserie 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   2760
         TabIndex        =   35
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label orlocal 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   1440
         TabIndex        =   34
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label idx 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Menu ajdu1 
      Caption         =   "&Add"
   End
   Begin VB.Menu f8443 
      Caption         =   "&Modifica"
   End
   Begin VB.Menu bo712 
      Caption         =   "&Borrar"
   End
   Begin VB.Menu fjh433 
      Caption         =   "&Zoom"
   End
   Begin VB.Menu fj8484 
      Caption         =   "&VerFormula"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu dki883 
      Caption         =   "Ver&ProduccionParte"
   End
   Begin VB.Menu djuer1 
      Caption         =   "&Reporte"
      Begin VB.Menu dk9893 
         Caption         =   "&0.GENERAL"
      End
      Begin VB.Menu mnuArchivoArray 
         Caption         =   "Novisible"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tordentd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txformxud  As New ADODB.Recordset

Dim txformxudf As New ADODB.Recordset

Private Sub ajdu1_Click()

    If Frame4.Visible = True Then Exit Sub
    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Nuevo"
    cmdGuardar.Enabled = True
    habilita 1
    producto.Enabled = True
    Frame5.Visible = False
    producto = ""
    producto.SetFocus

End Sub

Private Sub bo712_Click()

    Dim buf As String

    On Error GoTo cmd656_err

    If Frame4.Visible = True Then Exit Sub
    buf = "" & txformxud.Fields("producto")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If MsgBox("Desea Borra " + "" & txformxud.Fields("PRODUCTO"), 1, "Aviso") <> 1 Then
        Exit Sub

    End If

    txformxud.Delete
    Command1_Click

    Exit Sub
cmd656_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub bodega_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_bodega

    End If

End Sub

Private Sub buffer_DblClick()
    Command1_Click

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    Command1_Click

End Sub

Private Sub cmdAddEntry_Click()
    ajdu1_Click

End Sub

Private Sub cmdCerrar_Click()
    dlo132_Click

End Sub

Private Sub cmdDelete_Click()
    bo712_Click

End Sub

Private Sub cmdExit_Click()
    dlo132_Click

End Sub

Private Sub cmdHelp_Click()

End Sub

Private Sub cmdGuardar_Click()

    Dim found As Integer

    consulta_stock

    found = grabar()

End Sub

Private Sub cmdPrint_Click()

    'djuer1_Click
End Sub

Private Sub cmdSave_Click()
    f8443_Click

End Sub

Private Sub Command2_Click()
    carga_formula
    refresca_formula

End Sub

Private Sub Command3_Click()
    Frame4.Visible = False

End Sub

Private Sub Command4_Click()
    filtro

End Sub

Private Sub Command6_Click()
    refresca_formula

End Sub

Private Sub dbgrid13_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim found As Integer

    If KeyCode = 27 Then
        Text1.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "1" Then
            'If valida_receta(Trim("" & dbgrid13.columns("producto"))) = 0 Then
            '   MsgBox "No tiene receta ", 48, "Aviso"
            '   Exit Sub
            'End If
            producto = Trim("" & dbgrid13.columns("producto"))
            descripcio = Trim("" & dbgrid13.columns("descripcio"))
            unidad = Trim("" & dbgrid13.columns("unidad"))
            factor = Trim("" & dbgrid13.columns("factor"))
            unidadf = Trim("" & dbgrid13.columns("unidadf"))
            factorf = Trim("" & dbgrid13.columns("factorf"))
            formula = Trim("" & Trim("" & dbgrid13.columns("id")))
            consulta_stock
            Frame3.Visible = False

            'producto.SetFocus
        End If

        If opcion1 = "12" Then
            producto = Trim("" & dbgrid13.columns("producto"))
            descripcio = Trim("" & dbgrid13.columns("descripcio"))
            unidad = Trim("" & dbgrid13.columns("unidad"))
            factor = Trim("" & dbgrid13.columns("factor"))
            cantidad = Trim("" & dbgrid13.columns("cantidad"))
            formula = "" ' Trim("" & dbgrid13.columns("id"))
            Frame3.Visible = False
            producto.SetFocus

        End If

        If opcion1 = "2" Then
            bodega = Trim("" & dbgrid13.columns("codigo"))
            Frame3.Visible = False

            'producto.SetFocus
        End If

    End If

End Sub

Private Sub dk9893_Click()

    If Frame4.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    reporgen.NAMETABLA = "ordentrabajod"
    reporgen.Show 1

End Sub

Sub prueba_reporte()

    'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\formulacionesproducto.rpt", "")
End Sub

Private Sub Command1_Click()
    Frame1.Visible = True
    Frame1.Enabled = True
    'buffer = ""
    opcion1 = "1"
    ejecuta 1

End Sub

Sub ejecuta(sw As Integer)

    Dim cad As String

    If Len(buffer) = 0 Then
        cad = "SELECT * from ordentrabajod where ordentrabajo=" & idx

    End If

    If Len(buffer) > 0 Then
        cad = "SELECT *  from ordentrabajod   where ordentrabajo=" & idx & " and " & Combo1 & " like '" & buffer & "%'"

    End If

    If txformxud.State = 1 Then txformxud.Close
    txformxud.Open cad, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = txformxud
    dbGrid1.columns(0).Width = 4000
    dbGrid1.columns(1).Width = 2000

    If txformxud.RecordCount > 0 Then
        dbGrid1.SetFocus

    End If

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then

        'formulacion = dbGrid1.Columns(1)
        'Frame1.Visible = False
        'Frame1.Enabled = False
        'formulacion.SetFocus
        'formulacion_KeyPress 13
    End If

End Sub

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

Private Sub dki883_Click()

    If Frame4.Visible = True Then Exit Sub
    Frame4.Visible = True
    visualiza_parteproduccion

End Sub

Private Sub dlo132_Click()

    If Frame4.Visible = True Then
        Frame4.Visible = False
        Exit Sub

    End If

    If Frame3.Visible = True Then
        Frame3.Visible = False
        Exit Sub

    End If

    If Frame2.Visible = True Then
        habilita 0
        Frame2.Visible = False
        dbGrid1.Enabled = True
        Exit Sub

    End If

    tordentd.Hide
    Unload tordentd

End Sub

Private Sub f8443_Click()

    Dim buf As String

    On Error GoTo cmd456_err

    If Frame4.Visible = True Then Exit Sub
    buf = "" & txformxud.Fields("producto")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Modifica"
    Frame5.Visible = True
    cmdGuardar.Enabled = True
    pone_registro
    habilita 1
    'MsgBox "abc"
    producto.Enabled = False
    cantidad.SetFocus
    Exit Sub
cmd456_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fj8484_Click()

    If Frame4.Visible = True Then Exit Sub
    Frame4.Visible = True
    visualiza_receta

End Sub

Private Sub fjh433_Click()

    Dim buf As String

    On Error GoTo cmd556_err

    If Frame4.Visible = True Then Exit Sub
    buf = "" & txformxud.Fields("producto")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Zoom"
    cmdGuardar.Enabled = False
    Frame5.Visible = True
    pone_registro
    habilita 1
    producto.Enabled = False
    'MsgBox "ABC"
    cantidad.SetFocus
    Exit Sub
cmd556_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Activate()
    'agregar_menus
    Command1_Click

End Sub

Private Sub Form_Load()
    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.ListIndex = 0

    'MsgBox "ABC"
End Sub

Sub inicializa()
    unidadf = ""
    factorf = ""
    producto = ""
    descripcio = ""
    unidad = ""
    factor = ""
    cantidad = "1"
    bodega = ""
    formula = ""

    'refresca_formula
End Sub

Sub pone_registro()
    unidadf = Trim("" & txformxud.Fields("unidadf"))
    factorf = Trim("" & txformxud.Fields("factorf"))
    producto = Trim("" & txformxud.Fields("producto"))
    descripcio = Trim("" & txformxud.Fields("descripcio"))
    unidad = Trim("" & txformxud.Fields("unidad"))
    factor = Trim("" & txformxud.Fields("factor"))
    cantidad = Trim("" & txformxud.Fields("cantidad"))
    bodega = Trim("" & txformxud.Fields("bodega"))
    formula = Trim("" & txformxud.Fields("formula"))
    refresca_formula

End Sub

Sub grabando()
    txformxud.Fields("producto") = Trim(producto)
    txformxud.Fields("descripcio") = Trim(descripcio)
    txformxud.Fields("unidad") = Trim(unidad)
    txformxud.Fields("factor") = Val(factor)
    txformxud.Fields("unidadf") = Trim(unidadf)
    txformxud.Fields("factorf") = Val(factorf)
    txformxud.Fields("cantidad") = Val(cantidad)
    txformxud.Fields("bodega") = Trim(bodega)
    txformxud.Fields("formula") = Trim(formula)
    txformxud.Fields("ordentrabajo") = Val(idx)

End Sub

Private Sub grba1_Click()

End Sub

Function grabar()

    Dim found  As Integer

    Dim rbusca As New ADODB.Recordset

    found = valida()

    If found = 0 Then
        MsgBox "Campos invalidos", 48, "Aviso"
        Exit Function

    End If

    If Frame2.Caption = "Nuevo" Then
        rbusca.Open "select producto from ordentrabajod where ordentrabajo=" & idx & " and producto='" & producto & "'", cn, adOpenStatic, adLockOptimistic

        If rbusca.RecordCount > 0 Then
            rbusca.Close
            MsgBox "Ya existe  ", 48, "Aviso"
            Exit Function

        End If

        txformxud.AddNew
        grabando
        txformxud.Update
        dlo132_Click
        Exit Function

    End If

    If Frame2.Caption = "Modifica" Then
        grabando
        txformxud.Update
        dlo132_Click
        Exit Function

    End If

End Function

Function valida()

    If Len(producto) = 0 Then
        producto.SetFocus
        Exit Function

    End If

    If Len(descripcio) = 0 Then
        descripcio.SetFocus
        Exit Function

    End If

    If Len(unidad) = 0 Then
        unidad.SetFocus
        Exit Function

    End If

    If Val(factor) <= 0 Then
        factor.SetFocus
        Exit Function

    End If

    If Val(cantidad) <= 0 Then
        cantidad.SetFocus
        Exit Function

    End If

    If Len(bodega) = 0 Then
        bodega.SetFocus
        Exit Function

    End If

    If Len(formula) = 0 Then
        'producto.SetFocus
        Exit Function

    End If

    consulta_stock
    'if Val(stock) < Val(cantidad) Then
    'MsgBox "No existe Stock "
    'End If

    valida = 1

End Function

Sub habilita(sw As Integer)

    If sw = 0 Then

        ajdu1.Enabled = True
        f8443.Enabled = True
        bo712.Enabled = True
        fjh433.Enabled = True
        djuer1.Enabled = True
        djuer1.Enabled = True
        Picture1.Enabled = True
        dbGrid1.Enabled = True
            
    End If

    If sw = 1 Then

        ajdu1.Enabled = False
        f8443.Enabled = False
        bo712.Enabled = False
        fjh433.Enabled = False
        djuer1.Enabled = False
        djuer1.Enabled = False
        Picture1.Enabled = False
        dbGrid1.Enabled = False
        dbGrid1.Enabled = False
           
    End If
      
End Sub

Sub agregar_menus()

    Dim I As Integer

    For I = 1 To mnuArchivoArray.count - 1
        Unload mnuArchivoArray(I)
    Next
     
    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from archivo where menu='formulacion' and   estado='S'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    Do

        If mytablex.EOF Then Exit Do
        Agregarm "" & mytablex.Fields("descripcio"), mnuArchivoArray
        mytablex.MoveNext
    Loop
    mytablex.Close

End Sub

Sub Agregarm(TextoDeMenu As String, QueMenu As Object)

    Dim indice As Integer

    'MsgBox QueMenu.count
    indice = QueMenu.count

    Load QueMenu(indice)

    QueMenu(indice).Caption = TextoDeMenu
    QueMenu(indice).Visible = True

End Sub

Private Sub Image1_Click()
    consulta_bodega

End Sub

Private Sub Image4_Click()
    consulta_producto
    Exit Sub

    'If Len(Trim(orlocal)) = 0 Or Len(Trim(orserie)) = 0 Or Len(Trim(ornumero)) = 0 Then
    '      consulta_producto
    '      Exit Sub
    '   End If
    '   consulta_pedido
End Sub

Private Sub Label8_Click()
    consulta_stock

End Sub

Sub mnuarchivoarray_click(Index As Integer)

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    buf = mnuArchivoArray(Index).Caption
    mytablex.Open "select * from archivo where menu='formulacion' and descripcio='" & Trim(buf) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close

    End If

    'busca el reporte
    buf = mytablex.Fields("archivo")
    mytablex.Close
    'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\" & buf, "")

End Sub

Sub consulta_pedido()
    Combo2.Clear
    Combo2.AddItem "formulacion.Descripcio"
    Combo2.AddItem "formulacion.Producto"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "12"
    Text1.SetFocus
    'Command4_Click

End Sub

Sub consulta_producto()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.AddItem "Producto"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "1"
    Text1.SetFocus
    Command4_Click

End Sub

Sub consulta_bodega()
    Combo2.Clear
    Combo2.AddItem "Nombre"
    Combo2.AddItem "Codigo"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "2"
    Text1.SetFocus
    Command4_Click

End Sub

Private Sub producto_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        'If Len(Trim(orlocal)) = 0 Or Len(Trim(orserie)) = 0 Or Len(Trim(ornumero)) = 0 Then
        consulta_producto
        Exit Sub

        'End If
        'consulta_pedido
        'buscar en el pedido de venta
    End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    Command4_Click

End Sub

Sub filtro()

    Dim mytablex As New ADODB.Recordset

    Dim cad      As String

    'MsgBox opcion1
    If opcion1 = "12" Then  'pedido producto
        If Len(Text1) = 0 Then
            cad = "select Descripcio,Producto,Unidad,Factor,Cantidad,Local,Tipo,Serie,Numero from dpedidov where local='" & orlocal & "' and serie='" & orserie & "' and numero='" & ornumero & "'"

        End If

        If Len(Text1) > 0 Then
            cad = "select Descripcio,Producto,Unidad,factor,Cantidad,Local,Tipo,Serie,Numero from dpedidov where local='" & orlocal & "' and serie='" & orserie & "' and numero='" & ornumero & "' and " & Combo2 & " like '" & Text1 & "%'"

        End If

        'MsgBox cad
        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        dbgrid13.columns(0).Width = 5000
        dbgrid13.columns(1).Width = 2000
        'dbgrid13.columns(2).Width = 1000
        'dbgrid13.columns(3).Width = 1000
           
    End If

    If opcion1 = "1" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select producto.Descripcio,producto.Producto,producto.Unidad,producto.Factor,formulacion.Id,formulacion.factor as factorf,formulacion.unidad as unidadf from formulacion inner join producto on formulacion.producto=producto.producto "

        End If

        If Len(Text1) > 0 Then
            cad = "select select producto.Descripcio,producto.Producto,producto.Unidad,producto.Factor,formulacion.Id,formulacion.factor as factorf,formulacion.unidad as unidadf from formulacion inner join producto on formulacion.producto=producto.producto and  " & Combo2 & " like '" & Text1 & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        dbgrid13.columns(0).Width = 5000
        dbgrid13.columns(1).Width = 2000
        'dbgrid13.columns(2).Width = 1000
        'dbgrid13.columns(3).Width = 1000
           
    End If
   
    If opcion1 = "2" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select Nombre,Codigo from bodega "

        End If

        If Len(Text1) > 0 Then
            cad = "select Nombre,Codigo from bodega where  " & Combo2 & " like '" & Text1 & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        dbgrid13.columns(0).Width = 5000
        dbgrid13.columns(1).Width = 2000
        'dbgrid13.columns(2).Width = 1000
        'dbgrid13.columns(3).Width = 1000
           
    End If

    If mytablex.RecordCount > 0 Then
        dbgrid13.SetFocus

    End If

    Exit Sub

End Sub

Sub visualiza_receta()

    On Error GoTo cmd9021_err

    Dim mytablex As New ADODB.Recordset

    nround = ""
    mytablex.Open "select Producto,Descripcio,Unidad,factor,Cantidad from componente where id=" & Val(txformxud.Fields("formula")), cn, adOpenStatic, adLockOptimistic
    Set dbgrid3.DataSource = mytablex
    dbgrid3.columns(0).Width = 1000
    dbgrid3.columns(1).Width = 4000
    Exit Sub
cmd9021_err:
    MsgBox "Aviso en Visualiza receta " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub visualiza_parteproduccion()

    On Error GoTo cmd90212_err

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    sdx = 0
    mytablex.Open "select Fecha,Numero,Bodega,Producto,Descripcio,Unidad,Factor,Cantidad from parteproducciond where ordentrabajo=" & Val(txformxud.Fields("ordentrabajo")), cn, adOpenStatic, adLockOptimistic
    Set dbgrid3.DataSource = mytablex
    dbgrid3.columns(0).Width = 1000
    dbgrid3.columns(1).Width = 1000
    dbgrid3.columns(2).Width = 1000
    dbgrid3.columns(3).Width = 1000
    dbgrid3.columns(4).Width = 4000
    dbgrid3.columns(5).Width = 1000
    dbgrid3.columns(6).Width = 1000
    dbgrid3.columns(7).Width = 1000
    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + Val("" & mytablex.Fields("cantidad"))
        mytablex.MoveNext
    Loop
    nround = "" & sdx

    Exit Sub
cmd90212_err:
    MsgBox "Aviso en Visualiza parte produccion " + error$, 48, "Aviso"
    Exit Sub

End Sub

Function valida_receta(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from formulacion where producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        valida_receta = 1

    End If

    mytablex.Close

End Function

Function consulta_stock() As Double

    Dim mytablex As New ADODB.Recordset

    stock = ""
    mytablex.Open "select * from almacen where local='01' and producto='" & Trim(producto) & "' and bodega='" & Trim(bodega) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        stock = "" & mytablex.Fields("saldo")
        consulta_stock = Val("" & mytablex.Fields("saldo"))

    End If

    mytablex.Close

End Function

Sub refresca_formula()

    Dim sdx As Double

    On Error GoTo cmd7856_err

    If txformxudf.State = 1 Then
        txformxudf.Close

    End If

    Set txformxudf = Nothing
    sdx = 0
    txformxudf.Open "select * from ordentrabajodf where ordentrabajod=" & Val("" & txformxud.Fields("ordentrabajod")), cn, adOpenStatic, adLockOptimistic
    Set dbgrid5.DataSource = txformxudf
    Do

        If txformxudf.EOF Then Exit Do
        sdx = sdx + Val("" & txformxudf.Fields("totalcosto"))
        txformxudf.MoveNext
    Loop
    total = Format(sdx, "0.00")
    Exit Sub
cmd7856_err:
    MsgBox "Aviso en Refresca Formula " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub carga_formula()

    Dim sdx      As Double

    Dim sdx1     As Double

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from componente where id=" & Val("" & txformxud.Fields("formula")), cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        MsgBox "No existe Componentes ", 48, "Aviso"
        mytablex.Close
        Exit Sub

    End If

    If txformxudf.State = 1 Then
        txformxudf.Close

    End If

    Set txformxudf = Nothing
    cn.Execute ("delete from ordentrabajodf where ordentrabajod=" & Val("" & txformxud.Fields("ordentrabajod")))
    txformxudf.Open "select * from ordentrabajodf where ordentrabajod=" & Val("" & txformxud.Fields("ordentrabajod")), cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        txformxudf.AddNew
        txformxudf.Fields("ordentrabajod") = Val("" & txformxud.Fields("ordentrabajod"))
        txformxudf.Fields("producto") = Trim("" & mytablex.Fields("producto"))
        txformxudf.Fields("descripcio") = Trim("" & mytablex.Fields("descripcio"))
        txformxudf.Fields("unidad") = Trim("" & mytablex.Fields("unidad"))
        txformxudf.Fields("factor") = Val("" & mytablex.Fields("factor"))
        sdx = Val("" & mytablex.Fields("cantidad")) / Val("" & txformxud.Fields("factorf"))
        sdx = Val(Format(sdx, "0.00"))
        txformxudf.Fields("cantidad") = Val("" & txformxud.Fields("cantidad")) * sdx
        txformxudf.Fields("porcentajepeso") = Val("" & mytablex.Fields("porcentajepeso")) * sdx
        txformxudf.Fields("porcentajeMerma") = Val("" & mytablex.Fields("porcentajemerma")) * sdx
        sdx1 = Val(Format(Val("" & mytablex.Fields("costo")), "0.00"))
        txformxudf.Fields("costo") = sdx1 * sdx
        txformxudf.Fields("totalcosto") = Val("" & txformxudf.Fields("costo")) * Val("" & txformxudf.Fields("cantidad"))
        txformxudf.Fields("tipo") = Trim("" & mytablex.Fields("tipo"))
        txformxudf.Fields("explosion") = Trim("" & mytablex.Fields("explosion"))
        txformxudf.Fields("formula") = Val("" & mytablex.Fields("formula"))

        txformxudf.Fields("productof") = Trim("" & mytablex.Fields("productof"))
        'MsgBox "123"
        txformxudf.Update
        mytablex.MoveNext
    Loop
    mytablex.Close
    Exit Sub

End Sub

