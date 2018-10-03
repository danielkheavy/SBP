VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tfacipt 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Documentos Valorados"
   ClientHeight    =   8925
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   14985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Lista Precios"
      Height          =   4815
      Left            =   3480
      TabIndex        =   199
      Top             =   1800
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
         Picture         =   "tfacipt.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   201
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
         TabIndex        =   200
         Top             =   360
         Width           =   3375
      End
      Begin MSDBGrid.DBGrid DBGrid4 
         Height          =   3615
         Left            =   120
         OleObjectBlob   =   "tfacipt.frx":1212
         TabIndex        =   202
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
         TabIndex        =   203
         Top             =   360
         Width           =   135
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
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
      Height          =   8895
      Left            =   0
      TabIndex        =   189
      Top             =   0
      Visible         =   0   'False
      Width           =   14775
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
         TabIndex        =   195
         TabStop         =   0   'False
         Top             =   600
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
         TabIndex        =   194
         TabStop         =   0   'False
         Top             =   240
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   193
         TabStop         =   0   'False
         Top             =   240
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
         Height          =   375
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   192
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
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
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   191
         TabStop         =   0   'False
         Top             =   600
         Width           =   2895
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Habilitar Proveedor"
         Height          =   375
         Left            =   7200
         TabIndex        =   190
         Top             =   240
         Visible         =   0   'False
         Width           =   2535
      End
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7695
         Left            =   120
         TabIndex        =   196
         Top             =   960
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   13573
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
         TabIndex        =   197
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
   Begin VB.CheckBox sinigv 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sin Igv"
      Height          =   375
      Left            =   5400
      TabIndex        =   188
      Top             =   1920
      Width           =   975
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
      Left            =   10560
      MaxLength       =   60
      TabIndex        =   185
      TabStop         =   0   'False
      Top             =   1200
      Width           =   3375
   End
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
      Left            =   10560
      MaxLength       =   60
      TabIndex        =   184
      TabStop         =   0   'False
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox localf 
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
      Left            =   14880
      MaxLength       =   2
      TabIndex        =   170
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox local1 
      BackColor       =   &H00FFFFC0&
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
      Left            =   840
      MaxLength       =   2
      TabIndex        =   14
      Top             =   840
      Width           =   735
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "CodigoProveedor"
      Height          =   1935
      Left            =   3960
      TabIndex        =   152
      Top             =   2760
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox producto 
         Height          =   495
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   157
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
         Picture         =   "tfacipt.frx":2275
         Style           =   1  'Graphical
         TabIndex        =   156
         TabStop         =   0   'False
         ToolTipText     =   "Grabar registro"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox rcodigo 
         Height          =   495
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   154
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
         Picture         =   "tfacipt.frx":3487
         Style           =   1  'Graphical
         TabIndex        =   153
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
         TabIndex        =   158
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label32 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo Proveedor"
         Height          =   495
         Left            =   120
         TabIndex        =   155
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Datos Adicionales"
      Height          =   4095
      Left            =   2400
      TabIndex        =   147
      Top             =   1320
      Visible         =   0   'False
      Width           =   7215
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
         TabIndex        =   179
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
         MaxLength       =   11
         TabIndex        =   177
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
         TabIndex        =   174
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
         TabIndex        =   172
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
         Picture         =   "tfacipt.frx":4699
         Style           =   1  'Graphical
         TabIndex        =   151
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
         Picture         =   "tfacipt.frx":58AB
         Style           =   1  'Graphical
         TabIndex        =   150
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
         TabIndex        =   148
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label43 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "servicio  (D C *)"
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
         TabIndex        =   180
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label42 
         BackColor       =   &H00FFFF00&
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
         TabIndex        =   178
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label40 
         BackColor       =   &H00FFFF00&
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
         TabIndex        =   175
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label39 
         BackColor       =   &H00FFFF00&
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
         TabIndex        =   173
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFF00&
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
         TabIndex        =   149
         Top             =   1680
         Width           =   1695
      End
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   17760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   14925
      TabIndex        =   138
      Top             =   0
      Width           =   14985
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
         Picture         =   "tfacipt.frx":6ABD
         Style           =   1  'Graphical
         TabIndex        =   141
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
         Picture         =   "tfacipt.frx":7CCF
         Style           =   1  'Graphical
         TabIndex        =   140
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
         Picture         =   "tfacipt.frx":8EE1
         Style           =   1  'Graphical
         TabIndex        =   139
         TabStop         =   0   'False
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label znumero 
         Height          =   375
         Left            =   11040
         TabIndex        =   145
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label zserie 
         Height          =   375
         Left            =   10200
         TabIndex        =   144
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label ztipo 
         Height          =   375
         Left            =   9480
         TabIndex        =   143
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
         TabIndex        =   142
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Observaciones"
      Height          =   3855
      Left            =   1800
      TabIndex        =   130
      Top             =   1800
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
         Picture         =   "tfacipt.frx":A0F3
         Style           =   1  'Graphical
         TabIndex        =   136
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
         Picture         =   "tfacipt.frx":B305
         Style           =   1  'Graphical
         TabIndex        =   135
         ToolTipText     =   "Borrar registro"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox observa4 
         Height          =   375
         Left            =   120
         MaxLength       =   40
         TabIndex        =   134
         TabStop         =   0   'False
         Top             =   1800
         Width           =   5055
      End
      Begin VB.TextBox observa3 
         Height          =   375
         Left            =   120
         MaxLength       =   40
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   1440
         Width           =   5055
      End
      Begin VB.TextBox observa2 
         Height          =   375
         Left            =   120
         MaxLength       =   40
         TabIndex        =   132
         TabStop         =   0   'False
         Top             =   1080
         Width           =   5055
      End
      Begin VB.TextBox observa1 
         Height          =   375
         Left            =   120
         MaxLength       =   40
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   720
         Width           =   5055
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Recibos Pagos Adelantados"
      Height          =   3135
      Left            =   4440
      TabIndex        =   110
      Top             =   3720
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton Command9 
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
         Picture         =   "tfacipt.frx":C517
         Style           =   1  'Graphical
         TabIndex        =   129
         TabStop         =   0   'False
         ToolTipText     =   "Nuevo registro"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox acuenta 
         Height          =   375
         Left            =   2640
         MaxLength       =   6
         TabIndex        =   125
         Top             =   2040
         Visible         =   0   'False
         Width           =   1455
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
         Height          =   615
         Left            =   4320
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tfacipt.frx":D729
         Style           =   1  'Graphical
         TabIndex        =   124
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton cmdCancelar 
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
         Left            =   4320
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tfacipt.frx":DED7
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.TextBox renumero3 
         Height          =   375
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   121
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox renumero2 
         Height          =   375
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   114
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox renumero1 
         Height          =   375
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   113
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox retipo1 
         Height          =   375
         Left            =   120
         MaxLength       =   6
         TabIndex        =   112
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A Cuenta"
         Height          =   375
         Left            =   840
         TabIndex        =   126
         Top             =   2040
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label retotal3 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   122
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label retotal 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   120
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Adelantos"
         Height          =   375
         Left            =   840
         TabIndex        =   119
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label retotal2 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   118
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label retotal1 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   117
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero"
         Height          =   375
         Left            =   1200
         TabIndex        =   116
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         Height          =   375
         Left            =   2640
         TabIndex        =   115
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         Height          =   375
         Left            =   120
         TabIndex        =   111
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Cargar Productos "
      Height          =   2775
      Left            =   3960
      TabIndex        =   101
      Top             =   2760
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
         Picture         =   "tfacipt.frx":E685
         Style           =   1  'Graphical
         TabIndex        =   105
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
         Picture         =   "tfacipt.frx":F897
         Style           =   1  'Graphical
         TabIndex        =   104
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
         TabIndex        =   103
         Top             =   840
         Width           =   3495
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Chequear dia de Visita"
         Height          =   375
         Left            =   240
         TabIndex        =   102
         Top             =   1200
         Width           =   3495
      End
      Begin VB.Label Label24 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Carga"
         Height          =   375
         Left            =   240
         TabIndex        =   106
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ingreso de Lineas"
      Height          =   3255
      Left            =   2280
      TabIndex        =   59
      Top             =   1800
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
         Picture         =   "tfacipt.frx":10AA9
         Style           =   1  'Graphical
         TabIndex        =   77
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
         Picture         =   "tfacipt.frx":11CBB
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Grabar registro"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox t16 
         Height          =   375
         Left            =   6240
         MaxLength       =   2
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t15 
         Height          =   375
         Left            =   6240
         MaxLength       =   2
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t14 
         Height          =   375
         Left            =   6240
         MaxLength       =   2
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox t13 
         Height          =   375
         Left            =   6240
         MaxLength       =   2
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t12 
         Height          =   375
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t11 
         Height          =   375
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t10 
         Height          =   375
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox t9 
         Height          =   375
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t8 
         Height          =   375
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t7 
         Height          =   375
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t6 
         Height          =   375
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox t5 
         Height          =   375
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t4 
         Height          =   375
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t3 
         Height          =   375
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t2 
         Height          =   375
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox t1 
         Height          =   375
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.Label linea 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4440
         TabIndex        =   100
         Top             =   360
         Width           =   855
      End
      Begin VB.Label nt16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   99
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label nt15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   98
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   97
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label nt13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   96
         Top             =   840
         Width           =   735
      End
      Begin VB.Label nt12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   95
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label nt11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   94
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   93
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label nt9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   92
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label30 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   91
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label29 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   90
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label28 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   89
         Top             =   840
         Width           =   735
      End
      Begin VB.Label nt8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   88
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label nt7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   87
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   86
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label nt5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   85
         Top             =   840
         Width           =   735
      End
      Begin VB.Label nt4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   84
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label nt3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   83
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   82
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label nt1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   81
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tallas"
         Height          =   375
         Left            =   240
         TabIndex        =   80
         Top             =   840
         Width           =   975
      End
      Begin VB.Label nlinea 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   79
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Linea"
         Height          =   375
         Left            =   240
         TabIndex        =   78
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
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1560
      Width           =   615
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   17520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
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
      Left            =   17640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   17640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   1140
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
      Top             =   1560
      Width           =   735
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
      Left            =   4680
      MaxLength       =   3
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1920
      Width           =   615
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "tfacipt.frx":12ECD
      Height          =   5415
      Left            =   0
      OleObjectBlob   =   "tfacipt.frx":12EE1
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2400
      Width           =   14775
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
      Left            =   10560
      MaxLength       =   30
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox bodegaf 
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
      Left            =   12840
      MaxLength       =   2
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   10560
      MaxLength       =   2
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   840
      Width           =   735
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
      Left            =   7560
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
      Left            =   7560
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
      Left            =   7560
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
      Left            =   7560
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
      Height          =   375
      Left            =   4680
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
      Left            =   4680
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
      MaxLength       =   11
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1935
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
      Top             =   1200
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
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label precio 
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
      Left            =   0
      TabIndex        =   198
      Top             =   8520
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
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
      Left            =   9480
      TabIndex        =   187
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
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
      Left            =   9480
      TabIndex        =   186
      Top             =   1560
      Width           =   1095
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
      TabIndex        =   183
      Top             =   3360
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
      TabIndex        =   182
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label serieimp 
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
      TabIndex        =   181
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label41 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F8.VerCostos"
      Height          =   255
      Left            =   7200
      TabIndex        =   176
      Top             =   8640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label38 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Alm.Dest."
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
      Left            =   11280
      TabIndex        =   171
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label zlocal 
      Height          =   375
      Left            =   14880
      TabIndex        =   169
      Top             =   7920
      Width           =   855
   End
   Begin VB.Label Label37 
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   168
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label tflete 
      Alignment       =   1  'Right Justify
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
      Left            =   10080
      TabIndex        =   167
      Top             =   8640
      Width           =   1455
   End
   Begin VB.Label Label35 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Left            =   8640
      TabIndex        =   166
      Top             =   8640
      Width           =   1455
   End
   Begin VB.Label Label36 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Height          =   375
      Left            =   11520
      TabIndex        =   165
      Top             =   8640
      Width           =   1455
   End
   Begin VB.Label xtotper 
      Alignment       =   1  'Right Justify
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
      Left            =   13080
      TabIndex        =   164
      Top             =   8640
      Width           =   1695
   End
   Begin VB.Label txpercepcio 
      Alignment       =   1  'Right Justify
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
      Left            =   13080
      TabIndex        =   163
      Top             =   8280
      Width           =   1695
   End
   Begin VB.Label Label34 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Left            =   11520
      TabIndex        =   162
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label nbodega1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      Height          =   195
      Left            =   13920
      TabIndex        =   161
      Top             =   2640
      Width           =   45
   End
   Begin VB.Label NBODEGA 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      Height          =   195
      Left            =   13920
      TabIndex        =   160
      Top             =   2400
      Width           =   45
   End
   Begin VB.Label escompra 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6000
      TabIndex        =   159
      Top             =   7920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label cargado 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   14760
      TabIndex        =   146
      Top             =   7200
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label gravado 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7200
      TabIndex        =   137
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descuento"
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
      Left            =   8640
      TabIndex        =   128
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label adetotal 
      Alignment       =   1  'Right Justify
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
      Left            =   10080
      TabIndex        =   127
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label zona 
      AutoSize        =   -1  'True
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
      Left            =   6480
      TabIndex        =   109
      Top             =   8280
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
      TabIndex        =   108
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
      TabIndex        =   107
      Top             =   240
      Width           =   255
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   360
      Left            =   2760
      Picture         =   "tfacipt.frx":18C38
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
      TabIndex        =   58
      Top             =   9840
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
      TabIndex        =   57
      Top             =   9840
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
      TabIndex        =   56
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
      TabIndex        =   55
      Top             =   9600
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
      TabIndex        =   54
      Top             =   9600
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
      TabIndex        =   53
      Top             =   9600
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
      TabIndex        =   52
      Top             =   9600
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
      TabIndex        =   51
      Top             =   9600
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
      TabIndex        =   50
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
      TabIndex        =   49
      Top             =   9360
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
      TabIndex        =   48
      Top             =   9360
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
      TabIndex        =   47
      Top             =   9360
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
      Left            =   9720
      TabIndex        =   46
      Top             =   9360
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
      TabIndex        =   45
      Top             =   9360
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
      TabIndex        =   44
      Top             =   9360
      Width           =   495
   End
   Begin VB.Label flagcruce 
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
      Left            =   15120
      TabIndex        =   43
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label tipoclie 
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
      Left            =   15120
      TabIndex        =   42
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label flage 
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
      Left            =   15120
      TabIndex        =   41
      Top             =   960
      Width           =   255
   End
   Begin VB.Label txsubtotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   40
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label txdescuento 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   8640
      TabIndex        =   39
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label tximpuesto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   11520
      TabIndex        =   38
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label txneto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   7200
      TabIndex        =   37
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label acu 
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
      Left            =   15120
      TabIndex        =   36
      Top             =   720
      Width           =   255
   End
   Begin VB.Label ntcant 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
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
      Left            =   6480
      TabIndex        =   35
      Top             =   7920
      Width           =   615
   End
   Begin VB.Label txtotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   34
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
      TabIndex        =   33
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label17 
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
      Left            =   0
      TabIndex        =   32
      Top             =   7800
      Width           =   3135
   End
   Begin VB.Label presenta 
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
      Left            =   0
      TabIndex        =   31
      Top             =   8160
      Width           =   3135
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFC0&
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
      Left            =   9480
      TabIndex        =   30
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local.Dest"
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
      TabIndex        =   29
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Alm.Actual"
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
      Left            =   9480
      TabIndex        =   28
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   27
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   26
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFC0&
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
      Left            =   3600
      TabIndex        =   25
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   24
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   23
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFC0&
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
      Left            =   3600
      TabIndex        =   22
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
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
      Left            =   3600
      TabIndex        =   21
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
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
      Left            =   3600
      TabIndex        =   20
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   19
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label tipo 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local"
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
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   17
      Top             =   1560
      Width           =   855
   End
   Begin VB.Menu dnu834 
      Caption         =   "&Nuevo"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu grba1 
      Caption         =   "&Grabar"
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tfacipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bk2 As Variant
Dim xproducto As String
Dim opcion7 As Integer
Private Type campo_precio
    unidad As String
    factor As String
    precio As String
    costo As String
    margen As String
    stock As String
End Type
Dim c1 As String
Dim c2 As String
Dim c3 As String
Dim c4 As String
Dim c5 As String
Dim c6 As String
Dim c7 As String
Dim c8 As String
Dim c9 As String

Dim campo_precios(12) As campo_precio

Private Sub acuenta_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
Frame6.Visible = False
suma_retotal
fpago.SetFocus
End Sub

Private Sub acuenta_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   renumero3.SetFocus
   Exit Sub
End If

End Sub


Private Sub bo712_Click()

End Sub

Private Sub bodega_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(bodega) = 0 Then
   consulta_almacen
   Exit Sub
End If
found = busca_bodega("" & bodega, 0)
If found = 0 Then
   bodega = ""
   Exit Sub
End If
If bodegaf.Visible = True Then
   bodegaf.SetFocus
   Exit Sub
End If
partida.SetFocus
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
If KeyAscii <> 13 Then Exit Sub
If ttipo = "Z" Then
   If Len(bodegaf) = 0 Then
      bodegaf.SetFocus
      Exit Sub
   End If
   found = busca_bodega("" & bodegaf, 1)
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
Command1_Click

End Sub

Private Sub buffer_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> 13 And KeyCode <> 27 Then
   ejecuta 0
End If
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
Frame6.Visible = False
fpago.SetFocus
End Sub

Private Sub cmdDelete_Click()

End Sub

Private Sub cmdExit_Click()
dlo132_Click

End Sub

Private Sub cmdGrabar_Click()
Frame6.Visible = False
suma_retotal
fpago.SetFocus
End Sub


Private Sub cmdPrint_Click()

End Sub

Private Sub cmdHelp_Click()

End Sub

Private Sub cmdSave_Click()
grba1_Click
End Sub

Private Sub cmdSort_Click()

End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(codigo) = 0 Then
   consulta_codigo
   Exit Sub
End If
found = busca_codigo()
If found = 0 Then Exit Sub
If bodegaf.Visible = True Then
   localf = codigo
End If
fecha.SetFocus
End Sub

Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_codigo
End If
If KeyCode = &H26 Then
   If numero.Enabled = True Then
      numero.SetFocus
   End If
   Exit Sub
End If
If KeyCode = &H76 Then  'f7
   If tipoclie <> "C" And tipoclie <> "P" Then
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
DBGrid4.Refresh
tproducto = xproducto
carga_dbgrid4
End If

End Sub

Private Sub Command1_Click()
ejecuta 1
End Sub
Sub ejecuta(sw As Integer)
Dim buf As String
Dim buf1 As String
Dim buf2 As String
Dim xbuf As String
Dim rconsulta As New adodb.Recordset

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
If tipoclie = "I" Then
   buf2 = "tlocal"
End If
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
      If Len(buffer) = 0 Then
      buf = "select Tipo,Numero,Fecha,Total,Moneda as M from recibo where usado<>'S' and tipo='" & retipo1 & "' and codigo='" & codigo & "'"
      Else
      buf = "select Tipo,Numero,Fecha,Total,Moneda as M from recibo where usado<>'S' and tipo='" & retipo1 & "' and codigo='" & codigo & "' and " & Combo1 & " like '" & buffer & "%'"
      End If
   End If
If opcion1 = "1" Then
      xbuf = " tipodoc='" & acu & "'"
      If acu = "V" Then
         xbuf = " (tipodoc='A' or tipodoc='B' or tipodoc='C' or tipodoc='G' or tipodoc='D' )"
      End If
      If acu = "C" Then
         xbuf = " (tipodoc='J' or tipodoc='K' or tipodoc='L' or tipodoc='M' or tipodoc='P')"
      End If
      If Len(buffer) = 0 Then
         buf = "select Descripcio,Tipo from Tipo where " & xbuf
      Else
         buf = "select Descripcio,Tipo from tipo where " & xbuf & " and " & Combo1 & " like '" & buffer & "%'"
      End If
End If
If opcion1 = "2" Then
      If Len(buffer) = 0 Then
      buf = "select Nombre,Codigo,Codigo1 from  " & buf2
      Else
      buf = "select Nombre,Codigo,Codigo1 from " & buf2 & " where " & Combo1 & " like '" & buffer & "%'"
      End If
End If
If opcion1 = "3" Then
      If Len(buffer) = 0 Then
      buf = "select Nombre,Codigo from Vendedor "
      Else
      buf = "select Nombre,Codigo from Vendedor where " & Combo1 & " like '" & buffer & "%'"
      End If
End If
If opcion1 = "4" Then
      If Len(buffer) = 0 Then
      buf = "select Nombre,Codigo from Transpor "
      Else
      buf = "select Nombre,Codigo from Transpor where " & Combo1 & " like '" & buffer & "%'"
      End If
End If
  
If opcion1 = "5" Then
      If Len(buffer) = 0 Then
      buf = "select Descripcio,Fpago from Fpago where moneda='" & moneda & "'"
      Else
      buf = "select Descripcio,Fpago from Fpago where " & Combo1 & " like '" & buffer & "%' and moneda='" & moneda & "'"
      End If
End If
If opcion1 = "6" Then
      If Len(buffer) = 0 Then
      buf = "select Nombre,Codigo from Bodega "
      Else
      buf = "select Nombre,Codigo from Bodega where " & Combo1 & " like '" & buffer & "%'"
      End If
End If
If opcion1 = "7" Then
      If Len(buffer) = 0 Then
      buf = "select Nombre,Codigo from Bodega "
      Else
      buf = "select Nombre,Codigo from Bodega where " & Combo1 & " like '" & buffer & "%'"
      End If
End If


If opcion1 = "8" Or opcion1 = "50" Then
      If Len(buffer) = 0 Then
      buf = "select Producto.Descripcio,Producto.producto,Producto.Marca,precios.Unidad1 as Und1,precios.Factor1 as F,precios.pventa1 as Precio,Producto.Monedav as M,Producto.Familia as Fam,Producto.Subfamilia as Subfam,Producto.barras,producto.Igv from producto  left join precios on producto.producto=precios.producto  where precios.local='" & local1 & "'"
      Else
      buf = "select Producto.Descripcio,Producto.producto,Producto.Marca,precios.Unidad1 as Und1,precios.Factor1 as F,Precios.pVenta1 as Precio,Producto.Monedav as M,Producto.Familia as Fam,Producto.Subfamilia as Subfam,Producto.barras,producto.Igv from producto left join precios on producto.producto=precios.producto WHERE  precios.local='" & local1 & "'  and "
      buf = buf & Combo1 & " like '" & buffer & "%'"
      End If
End If
'---------------------------
      
'---------------------------
If opcion1 = "45" Then  'son compras a proveedores
If Len(buffer) = 0 Then
  buf = "select Producto.descripcio,Producto.producto,producto.marca,producto.unidad as Und1,producto.Factor as F,Producto.Costou as Precio,producto.monedac as M,producto.familia,producto.Subfamilia,codprov.codigo from codprov left join producto on codprov.producto=producto.producto where codprov.codigo='" & codigo & "'"
  Else
  buf = "select Producto.descripcio,Producto.producto,producto.marca,producto.unidad as Und1,producto.Factor as F,Producto.Costou as Precio,producto.monedac as M,producto.familia,producto.Subfamilia,codprov.codigo from codprov left join producto on codprov.producto=producto.producto where codprov.codigo='" & codigo & "' and  descripcio like '" & buffer & "%'"
End If
End If
If Combo2 <> "%" Then
   buf = buf & " and " & Combo2 & " like '" & buffer1 & "'"
End If
'MsgBox buf
If rconsulta.State = 1 Then rconsulta.Close
   'MsgBox buf
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
   
   Set dbGrid1.DataSource = rconsulta
   sw_consulta = 1
   
               If opcion1 = "444" Or opcion1 = "443" Or opcion1 = "21" Or opcion1 = "1" Or opcion1 = "2" Or opcion1 = "3" Or opcion1 = "4" Or opcion1 = "5" Or opcion1 = "6" Or opcion1 = "7" Then
               dbGrid1.columns(0).Width = 4000
               dbGrid1.columns(1).Width = 2000
               End If
               If opcion1 = "22" Or opcion1 = "23" Or opcion1 = "24" Then
               dbGrid1.columns(0).Width = 1000
               dbGrid1.columns(1).Width = 1500
               dbGrid1.columns(2).Width = 1500
               dbGrid1.columns(3).Width = 1500
               dbGrid1.columns(4).Width = 700
               End If
               
               If opcion1 = "8" Or opcion1 = "50" Or opcion1 = "45" Then
               dbGrid1.columns(0).Width = 5000
               dbGrid1.columns(1).Width = 1300
               dbGrid1.columns(2).Width = 1000
               dbGrid1.columns(3).Width = 900
               dbGrid1.columns(4).Width = 500
               dbGrid1.columns(5).Width = 900
               dbGrid1.columns(6).Width = 500
               dbGrid1.columns(7).Width = 800
               dbGrid1.columns(8).Width = 800
               dbGrid1.columns(9).Width = 1700
               'dbGrid1.Columns(10).Width = 500
               End If
If sw = 1 Then
               dbGrid1.SetFocus
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
If servicio <> "*" And servicio <> "D" And servicio <> "C" Then
   servicio = "*"
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
If servicio <> "*" And servicio <> "D" And servicio <> "C" Then
   servicio = "*"
End If
dlo132_Click
End Sub

Private Sub Command12_Click()
Frame8.Visible = False
DBGrid2.SetFocus
End Sub

Private Sub Command13_Click()
On Error GoTo cmd5665_err
Dim mytablex As New adodb.Recordset
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

Private Sub Command2_Click()
Dim sdx As Double
DBGrid2.columns("t1") = Val(t1)
DBGrid2.columns("t2") = Val(t2)
DBGrid2.columns("t3") = Val(t3)
DBGrid2.columns("t4") = Val(t4)
DBGrid2.columns("t5") = Val(t5)
DBGrid2.columns("t6") = Val(t6)
DBGrid2.columns("t7") = Val(t7)
DBGrid2.columns("t8") = Val(t8)
DBGrid2.columns("t9") = Val(t9)
DBGrid2.columns("t10") = Val(t10)
DBGrid2.columns("t11") = Val(t11)
DBGrid2.columns("t12") = Val(t12)
DBGrid2.columns("t13") = Val(t13)
DBGrid2.columns("t14") = Val(t14)
DBGrid2.columns("t15") = Val(t15)
DBGrid2.columns("t16") = Val(t16)
sdx = Val(t1) + Val(t2) + Val(t3) + Val(t4) + Val(t5) + Val(t6) + Val(t7) + Val(t8) + Val(t9) + Val(t10) + Val(t11) + Val(t12) + Val(t13) + Val(t14) + Val(t15) + Val(t16)
DBGrid2.columns("cantidad") = sdx
calcula_igv 0
Command3_Click
End Sub

Private Sub Command3_Click()
dlo132_Click
End Sub

Private Sub Command4_Click()
Dim sdx As Double
DBGrid2.columns("observa1") = "" & observa1
DBGrid2.columns("observa2") = "" & observa2
DBGrid2.columns("observa3") = "" & observa3
DBGrid2.columns("observa4") = "" & observa4
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
           DBGrid2.Col = 4
            DBGrid2.Row = DBGrid2.VisibleRows - 2

'DBGrid2.Col = 3
DBGrid2.SetFocus
End Sub

Private Sub Command9_Click()
acuenta = ""
retipo1 = ""
renumero1 = ""
renumero2 = ""
renumero3 = ""
retotal = ""
retotal1 = ""
retotal2 = ""
retotal3 = ""
suma_retotal
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim found As Integer
Dim buf As String
Dim xtemp As Variant
If KeyCode = &H70 Then  'f1
   If Len(dbGrid1.columns(0)) > 0 Then
      If opcion1 = "20" Then
         consulta_detalles
      End If
      Exit Sub
   End If
End If
If KeyCode = &H71 Then  'f2   cargar productos x bloque
   If Len(dbGrid1.columns(0)) > 0 Then
      If opcion1 = "8" Then
         consulta_bloques
      End If
      Exit Sub
   End If
End If
opcion3 = ""
If KeyCode = &H72 Then  'f3
   If Len(dbGrid1.columns(0)) > 0 Then
      If opcion1 = "8" Then
         opcion3 = "1"
         xproducto = "" & dbGrid1.columns(1)
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
   serie = Trim(dbGrid1.columns(1))
   numero = Trim(dbGrid1.columns(2))
   Frame1.Visible = False
   Frame1.Enabled = False
   numero.SetFocus
   numero_KeyPress 13
End If

If opcion1 = "21" Then
   retipo1 = Trim(dbGrid1.columns(1))
   Frame1.Visible = False
   Frame1.Enabled = False
   retipo1.SetFocus
   retipo1_KeyPress 13
End If
If opcion1 = "443" Then
   local1 = Trim(dbGrid1.columns(1))
   Frame1.Visible = False
   Frame1.Enabled = False
   local1.SetFocus
   local1_KeyPress 13
End If
If opcion1 = "444" Then
   localf = Trim(dbGrid1.columns(1))
   Frame1.Visible = False
   Frame1.Enabled = False
   localf.SetFocus
   localf_KeyPress 13
End If

If opcion1 = "22" Then
   renumero1 = Trim(dbGrid1.columns(1))
   retotal1 = Trim(dbGrid1.columns(3))
   suma_retotal
   Frame1.Visible = False
   Frame1.Enabled = False
   renumero1.SetFocus
   renumero1_KeyPress 13
End If
If opcion1 = "23" Then
   renumero2 = Trim(dbGrid1.columns(1))
   retotal2 = Trim(dbGrid1.columns(3))
   suma_retotal
   Frame1.Visible = False
   Frame1.Enabled = False
   renumero2.SetFocus
   renumero2_KeyPress 13
End If
If opcion1 = "24" Then
   renumero3 = Trim(dbGrid1.columns(1))
   retotal3 = Trim(dbGrid1.columns(3))
   suma_retotal
   Frame1.Visible = False
   Frame1.Enabled = False
   renumero3.SetFocus
   renumero3_KeyPress 13
End If
If opcion1 = "1" Then
   ttipo = Trim(dbGrid1.columns(1))
   Frame1.Visible = False
   Frame1.Enabled = False
   ttipo.SetFocus
   ttipo_KeyPress 13
End If

If opcion1 = "2" Then
   codigo = Trim(dbGrid1.columns(1))
   Frame1.Visible = False
   Frame1.Enabled = False
   codigo.SetFocus
   codigo_KeyPress 13
End If
If opcion1 = "3" Then
   vendedor = Trim(dbGrid1.columns(1))
   Frame1.Visible = False
   Frame1.Enabled = False
   vendedor.SetFocus
   vendedor_KeyPress 13
End If
If opcion1 = "4" Then
   transporte = Trim(dbGrid1.columns(1))
   Frame1.Visible = False
   Frame1.Enabled = False
   transporte.SetFocus
   transporte_KeyPress 13
End If
If opcion1 = "5" Then
   fpago = Trim(dbGrid1.columns(1))
   Frame1.Visible = False
   Frame1.Enabled = False
   fpago.SetFocus
   fpago_KeyPress 13
End If
If opcion1 = "6" Then
   bodega = Trim(dbGrid1.columns(1))
   Frame1.Visible = False
   Frame1.Enabled = False
   bodega.SetFocus
   bodega_KeyPress 13
End If
If opcion1 = "7" Then
   bodegaf = Trim(dbGrid1.columns(1))
   Frame1.Visible = False
   Frame1.Enabled = False
   bodegaf.SetFocus
   bodegaf_KeyPress 13
End If
If opcion1 = "50" Then
   producto = Trim(dbGrid1.columns(1))
   Frame1.Visible = False
   Frame1.Enabled = False
   producto.SetFocus
   producto_KeyPress 13
End If

If opcion1 = "8" Or opcion1 = "45" Then
   '------------------------
   
   '------------------------

   If Len("" & DBGrid2.columns("producto")) = 0 And Len(Trim("" & dbGrid1.columns(1))) > 0 Then
      found = verifica_doble(Trim("" & dbGrid1.columns(1)))
      If found = 1 Then
         MsgBox "Producto ya seleccionado", 48, "Aviso"
         dbGrid1.SetFocus
         Exit Sub
      End If
      'MsgBox ""
      
      xtemp = DBGrid2.Row
      'Data2.Refresh
      DBGrid2.Refresh
      'solo_ir_ultimo
      DBGrid2.Enabled = True
      DBGrid2.SetFocus
      If xtemp = -1 Then
         xtemp = 0
      End If
      DBGrid2.Row = xtemp
      DBGrid2.Col = 0
      DBGrid2.columns("producto") = Trim("" & dbGrid1.columns(1))
      found = busca_producto(Trim("" & DBGrid2.columns("producto")), 0)
      If found = 0 Then
         Exit Sub
      End If
      Frame1.Visible = False
      Frame1.Enabled = False
      'sumar_detalle
      'DBGrid2.Row = DBGrid2.VisibleRows - 1
      DBGrid2.Col = 4
      DBGrid2.SetFocus
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
Dim sdx As Double
sdx = Val(retotal1) + Val(retotal2) + Val(retotal3)
retotal = Format(sdx, "0.00")
adetotal = Format(Val(retotal), "0.00")
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


Private Sub DBGrid2_AfterColEdit(ByVal ColIndex As Integer)
Select Case ColIndex
       Case 0
       Case 3
End Select
End Sub

Private Sub DBGrid2_AfterColUpdate(ByVal ColIndex As Integer)
Dim found As Integer
Dim sdx As Double

Select Case ColIndex
       Case 0
            'found = busca_producto("" & DBGrid2.Columns(0), 0)
            'If found = 0 Then
            '   MsgBox "No existe producto", 48, "Aviso"
            '   Exit Sub
            'End If
            sumar_detalle
            DBGrid2.Col = 4
            DBGrid2.Row = DBGrid2.VisibleRows - 2
            DBGrid2.SetFocus
       Case 1
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 2
            DBGrid2.SetFocus
       
       Case 4
            'sdx = Val("" & DBGrid2.Columns(3)) * Val("" & DBGrid2.Columns(5))
            'DBGrid2.Columns(9) = Val(Format(sdx, "0.00"))
            'DBGrid2.Columns(7) = Val(Format(sdx, "0.00"))
            'calcula_igv
            'ir_ultimo
            sumar_detalle
            DBGrid2.Col = 6
            DBGrid2.Row = DBGrid2.VisibleRows - 2
            DBGrid2.SetFocus
       Case 6
            'sdx = Val("" & DBGrid2.Columns(3)) * Val("" & DBGrid2.Columns(5))
            'DBGrid2.Columns(9) = Val(Format(sdx, "0.00"))
            'DBGrid2.Columns(7) = Val(Format(sdx, "0.00"))
            'calcula_igv
            sumar_detalle
            DBGrid2.Col = 7
            DBGrid2.Row = DBGrid2.VisibleRows - 2
            DBGrid2.SetFocus
       Case 7
            'sdx = Val("" & DBGrid2.Columns(3)) * Val("" & DBGrid2.Columns(5))
            'DBGrid2.Columns(9) = Val(Format(sdx, "0.00"))
            'DBGrid2.Columns(7) = Val(Format(sdx, "0.00"))
            'calcula_igv
            sumar_detalle
            DBGrid2.Col = 7
            DBGrid2.Row = DBGrid2.VisibleRows - 2
            DBGrid2.SetFocus
       Case 8
            'If Val("" & DBGrid2.Columns(3)) > 0 Then
            '   sdx = Val("" & DBGrid2.Columns(7)) / Val("" & DBGrid2.Columns(3))
            '   DBGrid2.Columns(5) = Val(Format(sdx, "0.00"))
            '   DBGrid2.Columns(9) = Val("" & DBGrid2.Columns(7))
            '   calcula_igv
               sumar_detalle
               DBGrid2.Col = 0
               DBGrid2.Row = DBGrid2.VisibleRows - 1
               DBGrid2.SetFocus
            'End If
End Select
End Sub

Private Sub DBGrid2_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
Dim found As Integer
If ColIndex > 8 Then
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

     If Len("" & DBGrid2.columns("producto")) > 0 Then  'si ya existe no cambiar
        Cancel = True
        Exit Sub
     End If
Case 1
     If Len("" & DBGrid2.columns("producto")) = 0 Then  '
        Cancel = True
        Exit Sub
     End If
     If Len("" & DBGrid2.columns("descripcio")) = 0 Then  '
        Cancel = True
        Exit Sub
     End If


     
Case 4
     If Len("" & DBGrid2.columns("producto")) = 0 Then  '
        Cancel = True
        Exit Sub
     End If
     If Len("" & DBGrid2.columns("linea")) > 0 Then  'ojo no se puede poner si es talla
        Cancel = True
        Exit Sub
     End If
Case 6, 8, 14, 7
     If Len("" & DBGrid2.columns("producto")) = 0 Then  '
        Cancel = True
        Exit Sub
     End If
     
End Select
End Sub

Private Sub DBGrid2_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim found As Integer
Dim sdx As Double


'if bandera=""
Select Case ColIndex
Case 0
     If Len(DBGrid2.columns("producto")) = 0 Then
        Cancel = True
        Exit Sub
     End If
     If Len(DBGrid2.columns("producto")) > 14 Then
        Cancel = True
        Exit Sub
     End If
     found = verifica_doble("" & DBGrid2.columns("producto"))
     If found = 1 Then
        Cancel = True
        MsgBox "Producto ya Seleccionado", 48, "Aviso"
        Exit Sub
     End If
     found = busca_producto("" & DBGrid2.columns("producto"), 0)
     If found = 0 Then
        Cancel = True
        'MsgBox "No existe producto", 48, "Aviso"
        If Mid$("" & DBGrid2.columns("producto"), 1, 1) <> "!" Then    'si es codigo de proveedor
           consulta_producto "" & DBGrid2.columns("producto")
        End If
        'DBGrid2.Columns = 3
        Exit Sub
     End If
Case 1
     If Len(DBGrid2.columns("producto")) = 0 Then
        Cancel = True
        Exit Sub
     End If
     If Len(DBGrid2.columns("descripcio")) = 0 Then
        Cancel = True
        Exit Sub
     End If

Case 4
     If Len(DBGrid2.columns("producto")) = 0 Then
        Cancel = True
        Exit Sub
     End If
     If Not IsNumeric("" & DBGrid2.columns("cantidad")) Then
        Cancel = True
        Exit Sub
     End If
     sdx = Val("" & DBGrid2.columns("cantidad")) * Val("" & DBGrid2.columns("precio"))
     DBGrid2.columns("total") = sdx
     calcula_igv 0
Case 6
     If Len(DBGrid2.columns("producto")) = 0 Then
        Cancel = True
        Exit Sub
     End If
     If Not IsNumeric(DBGrid2.columns("precio")) Then
        Cancel = True
        Exit Sub
     End If
     
     sdx = Val("" & DBGrid2.columns("precio"))
     If sinigv.Value = 1 Then
        sdx = Val("" & DBGrid2.columns("precio")) + Val("" & DBGrid2.columns("precio")) * Val("" & DBGrid2.columns("igv")) / 100
        DBGrid2.columns("precio") = sdx
     End If
     
     sdx = Val("" & DBGrid2.columns("cantidad")) * Val("" & DBGrid2.columns("precio"))
     DBGrid2.columns("total") = sdx
     calcula_igv 0
     
Case 7
     If Len(DBGrid2.columns("producto")) = 0 Then
        Cancel = True
        Exit Sub
     End If
     If Not IsNumeric(DBGrid2.columns("deslipo")) Then
        Cancel = True
        Exit Sub
     End If
     sdx = Val("" & DBGrid2.columns("cantidad")) * Val("" & DBGrid2.columns("precio"))
     DBGrid2.columns("total") = sdx
     calcula_igv 0
Case 8
     If Len(DBGrid2.columns("producto")) = 0 Then
        Cancel = True
        Exit Sub
     End If
     If Not IsNumeric(DBGrid2.columns("total")) Then
        Cancel = True
        Exit Sub
     End If
     If Val("" & DBGrid2.columns("cantidad")) = 0 Then
        Cancel = True
        Exit Sub
     End If
     sdx = Val("" & DBGrid2.columns("total"))
     If sinigv.Value = 1 Then
        sdx = Val("" & DBGrid2.columns("total")) + Val("" & DBGrid2.columns("total")) * Val("" & DBGrid2.columns("igv")) / 100
        DBGrid2.columns("total") = sdx
     End If
     
     sdx = Val("" & DBGrid2.columns("total")) / Val("" & DBGrid2.columns("cantidad"))
     DBGrid2.columns("precio") = sdx
     calcula_igv 0
Case 14
     If Len(DBGrid2.columns("producto")) = 0 Then
        Cancel = True
        Exit Sub
     End If
     If Not IsNumeric(DBGrid2.columns("neto")) Then
        Cancel = True
        Exit Sub
     End If
     If Val("" & DBGrid2.columns("cantidad")) = 0 Then
        Cancel = True
        Exit Sub
     End If
     calcula_sinigv
     'calcula_igv 1
    

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
   Select Case DBGrid2.Col
          Case 0
               If Len(DBGrid2.columns("producto")) > 0 Then
                DBGrid2.Col = 4
                Exit Sub
          End If
          Case 4
               If Val(DBGrid2.columns("cantidad")) > 0 Then
                DBGrid2.Col = 8
                Exit Sub
          End If
          Case 5
               If Val(DBGrid2.columns("cantidad")) > 0 Then
                DBGrid2.Col = 6
                Exit Sub
          End If
          Case 6
               If Val(DBGrid2.columns("precio")) > 0 Then
                DBGrid2.Col = 8
                Exit Sub
          End If
          Case 7
               If Val(DBGrid2.columns("precio")) > 0 Then
                DBGrid2.Col = 8
                Exit Sub
          End If
          Case 8
               If Val(DBGrid2.columns("total")) > 0 Then
                DBGrid2.Col = 0
                DBGrid2.Row = DBGrid2.VisibleRows - 1
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
Dim found As Integer
Dim kproducto As String
On Error GoTo cmd34_err
ver_presenta
If KeyCode = &H77 Then  'f1
   If Len(codigo) = 0 Then
      MsgBox "debe existir cliente", 48, "Aviso"
      codigo.SetFocus
      Exit Sub
   End If

   If Len(DBGrid2.columns("producto")) > 0 And DBGrid2.Col = 2 Then
      If Val("" & DBGrid2.columns("precio")) <= 0 Or Val("" & DBGrid2.columns("cantidad")) = 0 Then
         MsgBox "Deben existir costos y Cantidades Ingresados ", 48, "Aviso"
         DBGrid2.SetFocus
         Exit Sub
      End If
      kproducto = "" & DBGrid2.columns("producto")
      found = ver_cambio_precios(kproducto)
      DBGrid2.SetFocus
      Exit Sub
   End If
End If
If bandera = "Ver" Then Exit Sub
 If KeyCode = &H70 Then  'f1
   If Len(DBGrid2.columns("producto")) > 0 And DBGrid2.Col = 2 Then
      xproducto = "" & DBGrid2.columns("producto")
      tproducto = xproducto
      Combo4.ListIndex = 0
      carga_dbgrid4
      Exit Sub
   End If
End If
If KeyCode = &H72 Then  'f3   crea el codigo interno de cada proveedor
   If acu <> "C" Then Exit Sub
   Frame8.Visible = True
   producto = ""
   rcodigo = ""
   producto.SetFocus
   Exit Sub
End If
If KeyCode = &H76 Then  'f7
   If Len(Trim("" & DBGrid2.columns("producto"))) > 0 Then
      xprodet.producto = Trim("" & DBGrid2.columns("producto"))
   End If
   xprodet.Show 1
   DBGrid2.SetFocus
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
If DBGrid2.Row = -1 Then
        MsgBox "No hay ningún registro para eliminar", vbInformation
        Exit Sub
End If
If MsgBox("Se va a eliminar el registro : está seguro ", _
   vbExclamation + vbYesNo, "Eliminar") = vbYes Then
   Data2.Recordset.Delete
   If Data2.Recordset.EOF = True And Data2.Recordset.BOF = True Then
      Exit Sub
   End If
   ir_ultimo
   Data2.Refresh
   'DBGrid2.Refresh
   DBGrid2.Col = 0
   DBGrid2.Row = DBGrid2.VisibleRows - 1
   DBGrid2.SetFocus
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
   If Len(DBGrid2.columns("producto")) = 0 Then
      consulta_producto ""
   End If
End If
If KeyCode = &H71 Then  'f2
   If Len(DBGrid2.columns("producto")) > 0 And Len(DBGrid2.columns("linea")) > 0 Then
      ingreso_tallas "" & DBGrid2.columns("linea")
   End If
End If
If KeyCode = &H72 Then  'f3
   If Len(DBGrid2.columns("producto")) > 0 Then
      ingreso_locales
   End If
End If

'If KeyCode = &H2D Then  'insert
'If KeyCode = &H28 Then  'flecha abajo
If KeyCode = &H28 Then  'flecha abajo inserta una nueva
         Exit Sub
         If DBGrid2.Col = 0 Then
            ir_ultimo
            If Len(DBGrid2.columns("producto")) > 0 And Len(DBGrid2.columns("descripcio")) > 0 And Len(DBGrid2.columns("unidad")) > 0 And Len(DBGrid2.columns("cantidad")) > 0 And Len(DBGrid2.columns("factor")) > 0 And Len(DBGrid2.columns("precio")) > 0 Then
            'Data2.Recordset.AddNew
            'Data2.Recordset.Update
            End If
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus
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
      dbGrid1.SetFocus
      Exit Sub
   End If
   Command8_Click
   Exit Sub
End If
If KeyCode = 13 Then
   'MsgBox opcion1
   If opcion3 = "1" Then
      Frame5.Visible = False
      dbGrid1.SetFocus
      Exit Sub
   End If
   'If opcion1 = "8" Then
   'If Len("" & DBGrid4.Columns(0)) > 0 And Val("" & DBGrid4.Columns(1)) > 0 And Len("" & DBGrid4.Columns(2)) > 0 Then
      'Data2.Recordset.Edit
      'Data2.Recordset.Fields("unidad") = "" & DBGrid4.Columns(0)
      'Data2.Recordset.Fields("factor") = "" & DBGrid4.Columns(1)
      'Data2.Recordset.Fields("precio") = "" & DBGrid4.Columns(3)
      'Data2.Recordset.Update
      DBGrid2.columns("unidad") = "" & DBGrid4.columns(0)
      DBGrid2.columns("factor") = Val("" & DBGrid4.columns(1))
      DBGrid2.columns("precio") = Val("" & DBGrid4.columns(2))
      buf = tipo_costo("" & ttipo)
      Select Case buf
             Case "C"
             DBGrid2.columns("precio") = Val("" & DBGrid4.columns(3)) ' / Val("" & DBGrid4.Columns(1))
             Case "V"
             DBGrid2.columns("precio") = Val("" & DBGrid4.columns(2))
      End Select
      sdx = Val("" & DBGrid2.columns("cantidad")) * Val("" & DBGrid2.columns("precio")) '* Val("" & DBGrid2.Columns("factor"))
      DBGrid2.columns("total") = sdx
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

Private Sub DBGrid4_UnboundReadData(ByVal RowBuf As MSDBGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim dr As Integer
Dim row_num As Integer
Dim r As Integer
Dim rows_returned As Integer
If ReadPriorRows Then
        dr = -1
    Else
        dr = 1
    End If
    If IsNull(StartLocation) Then
        If ReadPriorRows Then
           row_num = RowBuf.RowCount - 1
           'row_num = 9
        Else
           row_num = 0
        End If
    Else
        row_num = CLng(StartLocation) + dr
    End If
    rows_returned = 0
    For r = 0 To RowBuf.RowCount - 1
        If row_num < 0 Or row_num > 9 Then Exit For
        RowBuf.Value(r, 0) = campo_precios(row_num).unidad
        RowBuf.Value(r, 1) = campo_precios(row_num).factor
        RowBuf.Value(r, 2) = campo_precios(row_num).precio
        RowBuf.Value(r, 3) = campo_precios(row_num).costo
        RowBuf.Value(r, 4) = campo_precios(row_num).margen
        RowBuf.Value(r, 5) = campo_precios(row_num).stock
        RowBuf.Bookmark(r) = row_num
        row_num = row_num + dr
        rows_returned = rows_returned + 1
   Next r
   RowBuf.RowCount = rows_returned
End Sub

Private Sub destino_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
observa.SetFocus
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


Private Sub dlo132_Click()
Dim found As Integer
On Error GoTo cmd891_err
If Frame7.Visible = True Then
   Frame7.Visible = False
   fechae.SetFocus
   Exit Sub
End If
If Frame4.Visible = True Then
   Frame4.Visible = False
   dbGrid1.SetFocus
   Exit Sub
End If
If dbgrid3.Visible = True Then
   cerrar_dbgrid3
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
      Frame1.Visible = False
      Frame1.Enabled = False
      retipo1.SetFocus
      Exit Sub
   End If
   If opcion1 = "22" Then
      Frame1.Visible = False
      Frame1.Enabled = False
      renumero1.SetFocus
      Exit Sub
   End If
   If opcion1 = "23" Then
      Frame1.Visible = False
      Frame1.Enabled = False
      renumero2.SetFocus
      Exit Sub
   End If
   If opcion1 = "24" Then
      Frame1.Visible = False
      Frame1.Enabled = False
      renumero3.SetFocus
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
   If opcion1 = "8" Or opcion1 = "45" Then
      Frame1.Visible = False
      Frame1.Enabled = False
      'DBGrid2.Bookmark = bk2
      DBGrid2.Enabled = True
      DBGrid2.SetFocus
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
If Frame6.Visible = True Then
   Frame6.Visible = False
   fpago.SetFocus
   Exit Sub
End If
If bandera = "Nuevo" Or bandera = "Modifica" Then
   found = salir_sin_grabar()
   If found = 0 Then
      Exit Sub
   End If
End If
tfacipt.Hide
Unload tfacipt
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
servicio.SetFocus
End Sub

Private Sub fechasunat_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   cajero.SetFocus
   Exit Sub
End If

End Sub

Private Sub Form_Activate()
Dim found As Integer
'If Len(caja) = 0 Then
'   caja = "00"
'End If
'local1 = glocal
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
      numero.SetFocus
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
      numero.SetFocus
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
      numero.SetFocus
   End If
   
   If acu = "Z" Then
      'local1 = "01"
      ttipo = "Z"
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
      codigo = "01"
      fpago = "1"
      numero.SetFocus
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
   numero = znumero
   found = busca_tipo(1)  'pone el acu
   found = busca_registro(1)
   If found = 0 Then
      MsgBox "No existe", 48, "Aviso"
   End If
   local1.Enabled = False
   ttipo.Enabled = False
   serie.Enabled = False
   numero.Enabled = False
   sql_detalle
   sumar_detalle
   codigo.SetFocus
   DBGrid2.AllowUpdate = False
End If
If bandera = "Modifica" Then
   inicializa
   habilita_numero 0
   habilita_cabeza 0
   habilita_detalle 0
   local1 = zlocal
   ttipo = ztipo
   serie = zserie
   numero = znumero
   found = busca_tipo(1)  'pone el acu
   found = busca_registro(1)
   If found = 0 Then
      MsgBox "No existe", 48, "Aviso"
   End If
   local1.Enabled = False
   ttipo.Enabled = False
   serie.Enabled = False
   numero.Enabled = False
   sql_detalle
   sumar_detalle
   codigo.SetFocus
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
End Sub
Sub inicializa()
On Error GoTo cmd3_err
tproducto = ""
precio = ""
servicio = "*"
cajero = ""
tflete = ""
xtotper = ""
txpercepcio = ""
NBODEGA = ""
fechasunat = ""
opcion7 = 0

Label17 = ""
presenta = ""
ttipo = ""
serie = ""
numero = ""
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
adetotal = ""
acuenta = ""
retipo1 = ""
renumero1 = ""
renumero2 = ""
renumero3 = ""
retotal = ""
retotal1 = ""
retotal2 = ""
retotal3 = ""
zona = ""
observa1 = ""
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
bodega = "01"
localf = ""
bodegaf = ""
observa = ""
estado = ""
caja = ""
'local1 = glocal
vendedor = gusuario

paridad = "" & busca_paridadg(0)
borrar_detalle_todo_registro
sql_detalle
Exit Sub
cmd3_err:
MsgBox "Error en inicializa" & error$, 48, "Aviso"
Exit Sub
End Sub
Function verificar_registro()
Dim found As Integer
Dim mytablex As New adodb.Recordset
   mytablex.Open "SELECT * FROM " & cgusuario & " where  local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & numero & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount > 0 Then  'si existe
      verificar_registro = 1
  End If
mytablex.Close


End Function
Function busca_registro(sw As Integer)
Dim found As Integer
Dim mytablex As New adodb.Recordset
   mytablex.Open "SELECT * FROM " & cgusuario & " where  local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & numero & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount > 0 Then  'si existe
      pone_registro mytablex
      busca_registro = 1
      If sw = 1 Then
         found = cargar_registrod()
      End If
   If sw = 2 Then
      If "" & mytablex.Fields("yausado") <> "1" Then  'sino esta usado modificar
      If "" & mytablex.Fields("estado") = "2" Then
         busca_registro = 2
         found = cargar_registrod()
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
Sub pone_registro(mytablex As adodb.Recordset)
Dim found As Integer
tipoimp = "" & mytablex.Fields("tipoimp")
serieimp = "" & mytablex.Fields("serieimp")
numeroimp = "" & mytablex.Fields("numeroimp")

caja = "" & mytablex.Fields("caja")
turno = "" & mytablex.Fields("turno")
servicio = "" & mytablex.Fields("servicio")
cajero = "" & mytablex.Fields("usuario")
local1 = "" & mytablex.Fields("local")
adetotal = "" & mytablex.Fields("adetotal")
acuenta = "" & mytablex.Fields("acuenta")
retipo1 = "" & mytablex.Fields("retipo1")
renumero1 = "" & mytablex.Fields("renumero1")
renumero2 = "" & mytablex.Fields("renumero2")
renumero3 = "" & mytablex.Fields("renumero3")
retotal = "" & mytablex.Fields("retotal")
retotal1 = "" & mytablex.Fields("retotal1")
retotal2 = "" & mytablex.Fields("retotal2")
retotal3 = "" & mytablex.Fields("retotal3")
'---
zona = "" & mytablex.Fields("zona")
ttipo = "" & mytablex.Fields("tipo")
serie = "" & mytablex.Fields("serie")
numero = "" & mytablex.Fields("numero")
codigo = "" & mytablex.Fields("codigo")
partida = "" & mytablex.Fields("partida")
destino = "" & mytablex.Fields("destino")
fecha = "" & mytablex.Fields("fecha")
fechasunat = "" & mytablex.Fields("fechasunat")
fechae = "" & mytablex.Fields("fechae")
moneda = "" & mytablex.Fields("moneda")
vendedor = "" & mytablex.Fields("vendedor")
servicio = "" & mytablex.Fields("servicio")
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
found = busca_codigo()
suma_retotal
End Sub
Sub grabando(mytablex As adodb.Recordset)
On Error GoTo cmd781_err
mytablex.Fields("caja") = caja
If Len(caja) = 0 Then
   mytablex.Fields("caja") = "00"
End If

mytablex.Fields("turno") = turno
mytablex.Fields("servicio") = servicio
mytablex.Fields("adetotal") = Val(adetotal)
mytablex.Fields("acuenta") = Val(acuenta)
mytablex.Fields("retipo1") = retipo1
mytablex.Fields("renumero1") = renumero1
mytablex.Fields("renumero2") = renumero2
mytablex.Fields("renumero3") = renumero3
mytablex.Fields("retotal1") = Val(retotal1)
mytablex.Fields("retotal2") = Val(retotal2)
mytablex.Fields("retotal3") = Val(retotal3)
mytablex.Fields("retotal") = Val(retotal)
mytablex.Fields("tflete") = Val(tflete)
mytablex.Fields("zona") = zona
mytablex.Fields("nombre") = Trim(Mid$("" & Label17, 1, 35))
mytablex.Fields("estado") = "2"
mytablex.Fields("yausado") = "0"
mytablex.Fields("tipoclie") = tipoclie
mytablex.Fields("tipo") = ttipo
mytablex.Fields("serie") = serie
mytablex.Fields("numero") = numero
mytablex.Fields("codigo") = codigo
mytablex.Fields("partida") = partida
mytablex.Fields("destino") = destino
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
mytablex.Fields("acu1") = "" & acu1
mytablex.Fields("flage") = "" & flage
mytablex.Fields("hora") = Format(Now, "hh:MM")
mytablex.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
mytablex.Fields("fechasunat") = Format(fechasunat, "dd/mm/yyyy")
mytablex.Fields("total") = Val("" & txtotal)
'mytablex.Fields("recibe") = Val("" & txtotal)
mytablex.Fields("descuento") = Val("" & txdescuento)
mytablex.Fields("neto") = Val("" & txneto)
mytablex.Fields("gravado") = Val("" & gravado)
mytablex.Fields("impuesto") = Val("" & tximpuesto)
mytablex.Fields("subtotal") = Val("" & txsubtotal)
mytablex.Fields("percepcion") = Val("" & txpercepcio)

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
Exit Sub
cmd781_err:
MsgBox "Aviso en grabando " + error$, 48, "Aviso"
Exit Sub

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
Dim i As Integer
Combo4.Clear
      For i = 1 To 9
          Combo4.AddItem Format(i, "00")
      Next i
      Combo4.ListIndex = 0
End Sub

Private Sub grba1_Click()
Dim found As Integer
Dim sdx As Double
Dim mytablex As New adodb.Recordset
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
'MsgBox "grba1"
sumar_detalle
If MsgBox("Desea Grabar", 1, "Aviso") <> 1 Then
   DBGrid2.SetFocus
   dnu834.Enabled = True
   Exit Sub
End If
If bandera = "Nuevo" Then  'adicionar
  If Len(numero) = 0 Then
      mytablex.Open "SELECT * FROM tipo where    tipo='" & ttipo & "'", cn, adOpenKeyset, adLockOptimistic
      If mytablex.RecordCount > 0 Then  'si existe
         sdx = Val("" & mytablex.Fields("numero")) + 1
         numero = "" & sdx
      End If
      mytablex.Close
  End If
akp:
      found = verificar_registro()
      If found = 1 Then
         sdx = Val(numero) + 1
         numero = "" & sdx
         GoTo akp
      End If
End If
If Not IsNumeric(numero) Then
   numero.SetFocus
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

Private Sub Image1_Click()
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


Private Sub Label1_Click()
cmdSort_Click
End Sub


Function grabar()
Dim rs As Recordset
Dim i As Integer
Dim pracu As String
Dim buf1 As String
Dim found As Integer
Dim mytablex As New adodb.Recordset
Dim mytabley As New adodb.Recordset
Dim mytablez As New adodb.Recordset
Dim mytablea As New adodb.Recordset
Dim mytableb As New adodb.Recordset
Dim mytablexy As New adodb.Recordset
Dim te As String
Dim ts As String
Dim xc1 As Double
Dim xc2 As Double
Dim xc3 As Double
Dim xc4 As Double

Dim fila As Integer
Dim sw As Integer
Dim xbuf As String
On Error GoTo cmd761_err
'graba cabecera
sw = 0
acu1 = busca_tipox("" & tipo1)
If racu = "Z" Then  'abrir base datos traslado
   mytableb.Open "SELECT * FROM detalle where local='" & local1 & "' and tipo='TS' and serie='" & serie & "' and numero='" & numero & "'", cn, adOpenKeyset, adLockOptimistic
End If
'MsgBox dgusuariog
'MsgBox cgusuario
   'MsgBox cgusuario
   xbuf = "SELECT * FROM " & cgusuario & " where  local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & numero & "'"
   
   If mytablex.State = 1 Then mytablex.Close
   mytablex.Open xbuf, cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount = 0 Then  'si existe
      mytablex.AddNew
      grabando mytablex
      mytablex.Update
      found = busca_tipo(7)   'graba  el numero
      graba_yausado_guia "0"
      grabar = 1
   Else
      'mytablex.Edit
      grabando mytablex
      'MsgBox "" & mytablex.Fields("estado")
      mytablex.Update
      graba_yausado_guia "0"
      grabar = 1
   End If
   mytablex.Close
   'MsgBox ""

'-----grabar credito
buf1 = busca_fpagoc("" & fpago)  'credito ,letra
'MsgBox ""
If buf1 = "C" Or buf1 = "G" Then
   If valida_flag("" & racu) = 1 Or valida_flag("" & racu) = 2 Then  'compras o ventas
      grabar_cuentaxc
   End If
End If
'MsgBox ""
'----desapues ver si hubo adelantos
'MsgBox ""
If Len(retipo1) > 0 Then
   If Len(renumero1) > 0 Then
   found = graba_adelantos("", "", retipo1, renumero1, "S")
   End If
   If Len(renumero2) > 0 Then
   found = graba_adelantos("", "", retipo1, renumero2, "S")
   End If
   If Len(renumero3) > 0 Then
   found = graba_adelantos("", "", retipo1, renumero3, "S")
   End If
End If
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
cn.Execute ("delete from " & dgusuariog & " where local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & numero & "'")

'ahora borramos en la base datos si es traslado
If racu = "Z" Then
   cn.Execute ("delete from detalle where local='" & local1 & "' and tipo='TE' and serie='" & serie & "' and numero='" & numero & "'")
   cn.Execute ("delete from detalle where local='" & local1 & "' and tipo='TS' and serie='" & serie & "' and numero='" & numero & "'")
End If
'MsgBox ""
'GRABANDO EN detalle
mytablexy.Open "SELECT * FROM " & dgusuariog & " where local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & numero & "'", cn, adOpenKeyset, adLockOptimistic
Data2.Refresh
Set rs = Data2.Recordset.Clone
Do
If rs.EOF Then Exit Do
mytablexy.AddNew

For i = 0 To rs.Fields.count - 1
mytablexy.Fields(i) = rs.Fields(i)
Next i

mytablexy.Fields("local") = "" & local1
mytablexy.Fields("tipo") = "" & ttipo
mytablexy.Fields("serie") = "" & serie
mytablexy.Fields("numero") = "" & numero
mytablexy.Fields("vendedor") = "" & vendedor
mytablexy.Fields("moneda") = "" & moneda
mytablexy.Fields("bodega") = "" & bodega
mytablexy.Fields("codigo") = "" & codigo
mytablexy.Fields("localf") = "" & localf
mytablexy.Fields("bodegaf") = "" & bodegaf
mytablexy.Fields("acu") = "" & racu
mytablexy.Fields("acu1") = "" & acu1
mytablexy.Fields("flage") = "" & flage
mytablexy.Fields("tipoclie") = tipoclie
mytablexy.Fields("usuario") = "" & gusuario
If Len(cajero) > 0 Then
   mytablexy.Fields("usuario") = "" & cajero
End If

mytablexy.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
'mytablexy.Fields("hora") = Format(Now, "hh:MM")
mytablexy.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
mytablexy.Fields("estado") = "2"
mytablexy.Fields("caja") = caja
If Len(caja) = 0 Then
   mytablexy.Fields("caja") = "00"
End If
mytablexy.Fields("turno") = turno
mytablexy.Fields("servicio") = servicio
mytablexy.Update
'----
If racu = "Z" Then  'traslado
mytableb.AddNew
For i = 0 To rs.Fields.count - 1
mytableb.Fields(i) = rs.Fields(i)
Next i
mytableb.Fields("local") = "" & local1
mytableb.Fields("tipo") = "TS"
mytableb.Fields("serie") = "" & serie
mytableb.Fields("numero") = "" & numero
mytableb.Fields("vendedor") = "" & vendedor
mytableb.Fields("moneda") = "" & moneda
mytableb.Fields("bodega") = "" & bodega
mytableb.Fields("localf") = "" '& localf
mytableb.Fields("bodegaf") = "" '& bodegaf
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
   
   For i = 0 To rs.Fields.count - 1
   mytableb.Fields(i) = rs.Fields(i)
   Next i
   
   mytableb.Fields("local") = "" & codigo '& Mid$(codigo, 1, 3)
   mytableb.Fields("tipo") = "TE" '& ttipo
   mytableb.Fields("serie") = "" & serie
   mytableb.Fields("numero") = "" & numero
   
   mytableb.Fields("vendedor") = "" & vendedor
   mytableb.Fields("moneda") = "" & moneda
   mytableb.Fields("bodega") = "" & bodegaf
   mytableb.Fields("localf") = "" '& localf
   mytableb.Fields("bodegaf") = "" '& bodegaf
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
   If Len(caja) = 0 Then
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
rs.MoveNext
Loop


found = valida_flag("" & racu)
If found = 0 Then
End If
If found = 1 Or found = 2 Then
'MsgBox "Hola"
   descarga_saldo local1, ttipo, serie, numero, 0, ""
End If
If found = 3 Then
   'MsgBox ""
   descarga_saldo local1, "TS", serie, numero, 0, "1"
   descarga_saldo localf, "TE", serie, numero, 0, "1" 'mytablea productos
End If
'MsgBox ""

If racu = "Z" Then
   mytableb.Close
End If

'mytablex.Close
'mytablea.Close
'mytablexy.Close
'mytabley.Close
Exit Function
cmd761_err:
MsgBox "Aviso en grabar " + error$, 48, "Aviso"
Exit Function
End Function
Sub descarga_saldo(xlocal As String, xtipo As String, xserie As String, xnumero As String, sw As Integer, tipoarchv As String)
Dim sdx As Double
Dim signo As Double
Dim buf As String
Dim found As Integer
Dim sww As Integer
Dim mytablefa As New adodb.Recordset
Dim mytablex As New adodb.Recordset
Dim mytabley As New adodb.Recordset
On Error GoTo cmd19_err
sww = 0
'AHORA HAY QUE VALIDAR QUE no existe ya cruzado el documento----
 mytablefa.Open "SELECT * FROM " & cgusuario & " where  local='" & xlocal & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'", cn, adOpenKeyset, adLockOptimistic
 If mytablefa.RecordCount > 0 Then  'si existe
   If Len(mytablefa.Fields("tipo1")) > 0 And Len(mytablefa.Fields("serie1")) > 0 And Len(mytablefa.Fields("numero1")) > 0 Then
        found = ve_descarga("" & mytablefa.Fields("tipo1"))
        If found = 1 Then
         sww = 1
        End If
   End If
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
             Case "S", "J", "K", "L", "M", "P", "E"
             signo = 1
             Case "T", "A", "B", "C", "D", "G", "N"
             signo = -1
      End Select
      'MsgBox signo
      If "" & mytablex.Fields("acu") = "P" Or "" & mytablex.Fields("acu") = "J" Or "" & mytablex.Fields("acu") = "K" Or "" & mytablex.Fields("acu") = "L" Or "" & mytablex.Fields("acu") = "M" Then 'compras varia el precios y costo
         graba_costos mytablex
      End If
      
      '-------------------------------------------------
'busden:
If sww = 0 Then
   If mytabley.State = 1 Then mytabley.Close
   mytabley.Open "select * from almacen where local='" & Trim("" & mytablex.Fields("local")) & "' and producto='" & Trim("" & mytablex.Fields("producto")) & "' and bodega='" & Trim("" & mytablex.Fields("bodega")) & "'", cn, adOpenDynamic, adLockOptimistic 'adOpenKeyset, adLockOptimistic
   'MsgBox mytabley.RecordCount
   If mytabley.RecordCount = 0 Then 'si existe
      'MsgBox ""
      mytabley.AddNew
      mytabley.Fields("local") = "" & mytablex.Fields("local")
      mytabley.Fields("producto") = "" & mytablex.Fields("producto")
      mytabley.Fields("bodega") = "" & mytablex.Fields("bodega")
      sdx = Val("" & mytabley.Fields("saldo")) + signo * Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
      'MsgBox sdx
      mytabley.Fields("saldo") = sdx
      mytabley.Update
      Else
      If sw = 0 Then
         'mytabley.Edit
         'MsgBox ""
         sdx = Val("" & mytabley.Fields("saldo")) + signo * Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
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
Sub graba_costos(mytablex As adodb.Recordset)
Dim mytablexx As New adodb.Recordset
Dim mytabley As New adodb.Recordset
Dim sdx3 As Double 'costo en una unidad del empaque
Dim sdx4 As Double
Dim sdx As Double
Dim coSmer As Double
Dim cossala As Double
Dim canstock As Double
Dim saldoant As Double
Dim asdx As Double
Dim bsdx As Double
On Error GoTo cmd23_err
'MsgBox "L" & mytablex.Fields("local") & " P" & mytablex.Fields("producto") & " B" & mytablex.Fields("bodega")
   saldoant = 0
   mytablexx.Open "SELECT * FROM almacen where local='" & "" & mytablex.Fields("local") & "' and producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & "" & mytablex.Fields("bodega") & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablexx.RecordCount > 0 Then 'si existe
      saldoant = Val("" & mytablexx.Fields("saldo"))
   End If
   mytablexx.Close
sdx3 = (Val("" & mytablex.Fields("total")) / (Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))))   'costo empaque aque unidad
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
   coSmer = sdx3 * Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
   cossala = (Val("" & mytabley.Fields("costop"))) * saldoant
   canstock = saldoant + Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
   
   If sdx4 = 0 And canstock > 0 Then
      sdx4 = (coSmer + cossala) / canstock
      sdx4 = Format(sdx4, "0.00000")
      sdx4 = sdx4    '* Val("" & mytablex.Fields("factor1"))
      sdx3 = sdx3    '* Val("" & mytablex.Fields("factor1"))
   Else
      sdx4 = sdx4   '* Val("" & mytablex.Fields("factor1"))
      sdx3 = sdx3   '* Val("" & mytablex.Fields("factor1"))
   End If
   asdx = Val(Format(Val("" & mytabley.Fields("costou")), "0.00"))
   bsdx = Val(Format(sdx3, "0.00"))
   If asdx > bsdx Or asdx < bsdx And bsdx > 0 Then
      mytabley.Fields("ok") = "F"
   End If
   
   If sdx4 > 0 Then
      mytabley.Fields("costop") = sdx4
   End If
   If sdx3 > 0 Then
      mytabley.Fields("costou") = sdx3
      If Val("" & mytabley.Fields("costoini")) = 0 Then
         mytabley.Fields("costoini") = sdx3
      End If
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
   local1.SetFocus
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

If Len(serie) = 0 Then
   serie.SetFocus
   Exit Function
End If
'If Len(numero) = 0 Then
'   numero.SetFocus
'   Exit Function
'End If


If bandera = "Nuevo" Then  'adicionar
   If Len(numero) > 0 Then
      found = verificar_registro()
      If found = 1 Then
         MsgBox "Modo adicion,Ya existe el numero,cambie por otro", 48, "Aviso"
         numero = ""
         numero.SetFocus
         Exit Function
      End If
   End If
End If

If Len(codigo) = 0 Then
   codigo.SetFocus
   Exit Function
End If


found = busca_codigo()
If found = 0 Then
   codigo.SetFocus
   Exit Function
End If
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
If Len(bodega) = 0 Then
   bodega.SetFocus
   Exit Function
End If
found = busca_bodega("" & bodega, 0)
If found = 0 Then
   bodega = ""
   Exit Function
End If
If bodegaf.Visible = True Then
   'If Len(localf) = 0 Then
   '   localf.SetFocus
   '   Exit Function
   'End If
   'found = busca_local1("" & localf)
   'If found = 0 Then
   '   localf = ""
   '   localf.SetFocus
   '   Exit Function
   'End If
   localf = codigo
   If Len(bodegaf) = 0 Then
      bodegaf.SetFocus
      Exit Function
   End If
   found = busca_bodega("" & bodegaf, 1)
   If found = 0 Then
      bodegaf = ""
      bodegaf.SetFocus
      Exit Function
   End If
End If
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

valida = 1
Exit Function
cmd1934_err:
MsgBox "Aviso en valida " + error$, 48, "Aviso"
Exit Function
End Function

Private Sub Label10_Click()
If codigo.Enabled = False Then Exit Sub
Frame6.Visible = True
retipo1.SetFocus
End Sub

Private Sub modif2_Click()
End Sub

Private Sub Label4_Click()
Dim found As Integer
found = leer_archivo_texto()
fecha.SetFocus
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
    If Len(numero) = 0 Then
      'found = busca_tipo(9)
      'If found = 0 Then
      '   numero.SetFocus
      '   Exit Sub
      'End If
    End If
    If Len(numero) > 0 Then
       found = verificar_registro()
       If found = 1 Then
          MsgBox "Modo adicion,Ya existe el numero,cambie por otro", 48, "Aviso"
          numero = ""
          numero.SetFocus
          Exit Sub
       End If
    End If
    codigo.SetFocus
    Exit Sub
End If
If Len(numero) = 0 Then
   numero.SetFocus
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

DBGrid2.Enabled = True
         sql_detalle
         DBGrid2.Row = DBGrid2.VisibleRows - 1
         DBGrid2.SetFocus
End Sub

Private Sub observa_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   destino.SetFocus
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
   observa1.SetFocus
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
   If bodegaf.Visible = True Then
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

Private Sub rcodigo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub renumero1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
KeyAscii = 0
Exit Sub
End If
suma_retotal
renumero2.SetFocus

End Sub

Private Sub renumero1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   retipo1.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   consulta_adelanto1
End If

End Sub

Private Sub renumero2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
KeyAscii = 0
Exit Sub
End If
suma_retotal
renumero3.SetFocus

End Sub

Private Sub renumero2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   renumero1.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   consulta_adelanto2
End If

End Sub

Private Sub renumero3_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
KeyAscii = 0
Exit Sub
End If
suma_retotal
Frame6.Visible = False
fpago.SetFocus

End Sub

Private Sub renumero3_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   renumero2.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   consulta_adelanto3
End If

End Sub

Private Sub retipo1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   Frame6.Visible = False
   fpago.SetFocus
   Exit Sub
End If
renumero1.SetFocus
End Sub

Private Sub retipo1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_retipo1
End If

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
numero.SetFocus
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

Private Sub Text2_Change()

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
fechasunat.SetFocus

End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)

End Sub

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

Private Sub ttipo_KeyPress(KeyAscii As Integer)
Dim found As Integer
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
Dim mytablex As New adodb.Recordset
Dim sdx As Double
'Label16 = ""
racu = ""

   mytablex.Open "SELECT * FROM tipo where   tipo='" & ttipo & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount = 0 Then
      mytablex.Close
      Exit Function
   End If
   
       If acu = "V" Or acu = "C" Then
         Select Case "" & mytablex.Fields("tipodoc")
                Case "A", "B", "C", "G", "D"
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
      If IsNumeric("" & numero) Then
         'mytablex.Edit
         mytablex.Fields("numero") = "" & numero
         mytablex.Update
      End If
      End If
      If sw = 9 Then
      sdx = Val("" & mytablex.Fields("numero")) + 1
      numero = "" & sdx
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
Dim mytablex As New adodb.Recordset
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
Dim mytablex As New adodb.Recordset
Label17 = ""
If tipoclie = "P" Then
mytablex.Open "SELECT * FROM proveedo where codigo='" & codigo & "'", cn, adOpenKeyset, adLockOptimistic
End If
If tipoclie = "C" Then
mytablex.Open "SELECT * FROM clientes where codigo='" & codigo & "'", cn, adOpenKeyset, adLockOptimistic
End If
If tipoclie = "I" Then
mytablex.Open "SELECT * FROM tlocal where codigo='" & codigo & "'", cn, adOpenKeyset, adLockOptimistic
End If
If mytablex.RecordCount = 0 Then 'si existe
      mytablex.Close
      Exit Function
End If

   Label17 = "" & mytablex.Fields("nombre")
   If Len(moneda) = 0 Then
      moneda = "" & mytablex.Fields("moneda")
   End If
   If tipoclie <> "I" Then
      
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
   
   busca_codigo = 1
mytablex.Close
End Function
Function busca_vendedor()
zona = ""
Dim rsexiste As New adodb.Recordset
   rsexiste.Open "SELECT * FROM vendedor where  codigo='" & vendedor & "'", cn, adOpenKeyset, adLockOptimistic
   If rsexiste.RecordCount > 0 Then  'si existe
      busca_vendedor = 1
      zona = "" & rsexiste.Fields("zona")
   End If
End Function
Function busca_local1(buf As String)
Dim rsexiste As New adodb.Recordset
   rsexiste.Open "SELECT * FROM tlocal where  codigo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic
   If rsexiste.RecordCount > 0 Then  'si existe
      busca_local1 = 1
   End If

End Function
Function busca_transporte()

Dim rsexiste As New adodb.Recordset
   rsexiste.Open "SELECT * FROM transpor where  codigo='" & transporte & "'", cn, adOpenKeyset, adLockOptimistic
   If rsexiste.RecordCount > 0 Then  'si existe
      busca_transporte = 1
   End If

End Function
Function busca_fpago()
Dim rsexiste As New adodb.Recordset
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
Function busca_bodega(buf As String, sw As Integer)
Dim mytablex As New adodb.Recordset
If sw = 0 Then
NBODEGA = ""
End If
If sw = 1 Then
nbodega1 = ""
End If


   mytablex.Open "SELECT * FROM bodega where  codigo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount > 0 Then  'si existe
      
   busca_bodega = 1
   If sw = 0 Then
      NBODEGA = Mid$("" & mytablex.Fields("nombre"), 1, 10)
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
               Data2.Refresh
               DBGrid2.Refresh
               'If Data2.Recordset.EOF = True And Data2.Recordset.BOF = True Then
               '   Data2.Recordset.AddNew
               '   Data2.Recordset.Update
               'End If
Exit Sub
cmd34_err:
MsgBox "Error en select " & error$, 48, "Aviso"
Exit Sub
End Sub
Function busca_producto(buf As String, sw As Integer)
Dim mytablex As New adodb.Recordset
Dim mytabley As New adodb.Recordset
Dim xbuf As String
Dim found As Integer
Dim sw1 As Integer
Dim ybuf As String
Dim buf1 As String
Dim i As Integer
Dim ssw As Integer

xbuf = buf
sw1 = 0
ybuf = ""
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
         '-- ahora busca los precios
         mytabley.Open "SELECT * FROM precios where  producto='" & "" & mytablex.Fields("producto") & "' and local='" & local1 & "'", cn, adOpenKeyset, adLockOptimistic
         If mytabley.RecordCount = 0 Then  'si existe
            MsgBox "No existe Precio venta en dicho Local ", 48, "Aviso"
            mytablex.Close
            mytabley.Close
            Exit Function
         End If
         graba_temporald mytablex, mytabley, sw
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
Sub graba_temporald(mytablex As adodb.Recordset, mytabley As adodb.Recordset, sw As Integer)
Dim found As Integer
Dim pventa1 As Double
Dim costou As Double
Dim buf As String
Dim mytables As New adodb.Recordset
pventa1 = Val("" & mytabley.Fields("pventa1"))
costou = Val("" & mytablex.Fields("costou"))
If "" & moneda = "S" Then
   If "" & mytablex.Fields("monedav") = "D" Then
      pventa1 = Val("" & mytabley.Fields("pventa1")) * Val("" & busca_paridadg(0))
   End If
   If "" & mytablex.Fields("monedaC") = "D" Then
      costou = Val("" & mytablex.Fields("costou")) * Val("" & busca_paridadg(0))
   End If
End If
If "" & moneda = "D" Then
   If "" & mytablex.Fields("monedav") = "S" Then
      pventa1 = Val("" & mytabley.Fields("pventa1")) / Val("" & busca_paridadg(0))
   End If
   If "" & mytablex.Fields("monedaC") = "S" Then
      costou = Val("" & mytablex.Fields("costou")) / Val("" & busca_paridadg(0))
   End If
End If


mytables.Open "SELECT * FROM DUENO where  local='" & local1 & "' and producto='" & "" & mytablex.Fields("producto") & "' ", cn, adOpenKeyset, adLockOptimistic
If mytables.RecordCount > 0 Then  'si existe
   DBGrid2.columns("ccosto") = Trim("" & mytables.Fields("codigo"))
End If
mytables.Close

DBGrid2.columns("producto") = "" & mytablex.Fields("producto")
'dbGrid2.Columns("proveedorp") = "" '& mytablex.Fields("proveedor1")
DBGrid2.columns("tipo") = "" & ttipo
DBGrid2.columns("serie") = "" & serie
DBGrid2.columns("numero") = "" & numero
DBGrid2.columns("vendedor") = "" & vendedor
DBGrid2.columns("descripcio") = Mid$("" & mytablex.Fields("descripcio"), 1, 60)
DBGrid2.columns("cantidad") = 1
DBGrid2.columns("unidad") = "" & mytabley.Fields("unidad1")
DBGrid2.columns("factor") = Val("" & mytabley.Fields("factor1"))
DBGrid2.columns("precio") = pventa1
DBGrid2.columns("total") = pventa1
DBGrid2.columns("subtotal") = pventa1
DBGrid2.columns("descuento") = Val("" & mytablex.Fields("isc"))
'DBGrid2.Columns(13) = Val("" & mytablex.Fields("tax"))
If valida_flag("" & racu) = "2" Then  'compras
DBGrid2.columns("unidad") = "" & mytablex.Fields("unidad")
DBGrid2.columns("factor") = Val("" & mytablex.Fields("factor"))
DBGrid2.columns("precio") = costou * Val("" & mytablex.Fields("factor"))
DBGrid2.columns("total") = costou * Val("" & mytablex.Fields("factor"))
DBGrid2.columns("subtotal") = costou * Val("" & mytablex.Fields("factor"))

End If

If valida_flag("" & racu) = "1" Then 'ventas
DBGrid2.columns("unidad") = "" & mytabley.Fields("unidad1")
DBGrid2.columns("factor") = Val("" & mytabley.Fields("factor1"))
DBGrid2.columns("precio") = pventa1
DBGrid2.columns("total") = pventa1
DBGrid2.columns("subtotal") = pventa1
End If

      buf = tipo_costo("" & ttipo)
      Select Case buf
             Case "V"
             DBGrid2.columns("precio") = pventa1
      End Select

DBGrid2.columns("deslipo") = 0
DBGrid2.columns("tax") = 0
DBGrid2.columns("flete") = Val("" & mytablex.Fields("flete"))
DBGrid2.columns("impuesto") = 0
DBGrid2.columns("igv") = Val("" & mytablex.Fields("igv"))
DBGrid2.columns("percepcion") = Val("" & mytablex.Fields("percepcion"))
DBGrid2.columns("linea") = "" & mytablex.Fields("linea")

DBGrid2.columns("descuento") = 0
DBGrid2.columns("neto") = 0

'---------pone a quien pertenece --------------------
DBGrid2.columns("l1") = "" '& mytablex.Fields("c11")
DBGrid2.columns("l2") = "" '& mytablex.Fields("c12")
DBGrid2.columns("l3") = "" '& mytablex.Fields("c13")
DBGrid2.columns("l4") = "" '& mytablex.Fields("c14")

'LAS FAMILIAS+SUBFAMILIA+MARCA+SECCION
DBGrid2.columns("familia") = "" & mytablex.Fields("Familia")
DBGrid2.columns("subfamilia") = "" & mytablex.Fields("subFamilia")
DBGrid2.columns("marca") = "" & mytablex.Fields("marca")
DBGrid2.columns("hora") = Format(Now, "hh:MM:ss")
'If bodega = "01" Then
'   found = ver_docena1(mytabley)
'End If
If racu = "E" Or racu = "N" Then  'si es nota credito por compras o ventas
   If Val("" & DBGrid2.columns("precio")) >= 0 Then
      DBGrid2.columns("precio") = -Val("" & DBGrid2.columns("precio"))
   End If
End If

'-----------------------------
calcula_igv 0
End Sub

Sub suma_linea()
Dim sdx As Double
'sdx = Val("" & Data2.Recordset.Fields("precio")) * Val("" & Data2.Recordset.Fields("cantidad"))
'Data2.Recordset.Fields("total") = Val(Format(sdx, "0.00"))
'Data2.Recordset.Fields("neto") = Val(Format(sdx, "0.00"))
End Sub
Sub calcula_igv(sw As Integer)
Dim sdx As Double
Dim sdx1 As Double
Dim sdx2 As Double
Dim tdscto As Double
Dim tdscto1 As Double
Dim found As Integer
If racu = "E" Or racu = "N" Then  'si es nota credito por compras o ventas
   If Val("" & DBGrid2.columns("precio")) >= 0 Then
      DBGrid2.columns("precio") = -Val("" & DBGrid2.columns("precio"))
      DBGrid2.columns("total") = Val("" & DBGrid2.columns("precio")) * Val("" & DBGrid2.columns("cantidad"))
      
   End If
End If
tdscto = Val("" & DBGrid2.columns("total")) * Val("" & DBGrid2.columns("deslipo")) / 100       'calcular descuento
DBGrid2.columns("descuento") = tdscto  'total descuento
DBGrid2.columns("total") = Val("" & DBGrid2.columns("total")) - Val("" & DBGrid2.columns("descuento")) 'cobrar
DBGrid2.columns("subtotal") = Val("" & DBGrid2.columns("total")) 'subtotal
DBGrid2.columns("impuesto") = 0
DBGrid2.columns("neto") = Val("" & DBGrid2.columns("subtotal")) + Val("" & DBGrid2.columns("descuento"))
If Val("" & DBGrid2.columns("total")) > 0 And Val("" & DBGrid2.columns("igv")) > 0 Then
   sdx2 = 1 + Val("" & DBGrid2.columns("igv")) / 100
   sdx1 = Val(DBGrid2.columns("total")) / sdx2
   DBGrid2.columns("subtotal") = sdx1  'subtotal
   sdx = Val("" & DBGrid2.columns("total")) - Val("" & DBGrid2.columns("subtotal"))
   DBGrid2.columns("impuesto") = sdx  'impuesto
   DBGrid2.columns("descuento") = tdscto
   DBGrid2.columns("neto") = Val("" & DBGrid2.columns("subtotal")) + Val("" & DBGrid2.columns("descuento"))
End If
DBGrid2.columns("tpercepcio") = Val(Format(Val("" & DBGrid2.columns("total")) * Val("" & DBGrid2.columns("percepcion")) / 100, "0.00"))
'PERCEPCIONÇ
'Data1.Recordset.Fields("total_flet") = Val(Format(xtotal * Val("" & Data1.Recordset.Fields("precio_fle")) / 100, "0.00"))
End Sub
Sub calcula_sinigv()
Dim sdx As Double
Dim sdx1 As Double
Dim sdx2 As Double
Dim found As Integer
'debe sumar el igv
'DBGrid2.Columns(13) = Val("" & DBGrid2.Columns(3)) * Val("" & DBGrid2.Columns(5))
If Val("" & DBGrid2.columns("cantidad")) > 0 And Val("" & DBGrid2.columns("neto")) > 0 Then
   sdx = Val("" & DBGrid2.columns("neto")) * Val("" & DBGrid2.columns("deslipo")) / 100 'descuento
   DBGrid2.columns("descuento") = sdx  'descuento
   DBGrid2.columns("subtotal") = Val("" & DBGrid2.columns("neto")) - Val("" & DBGrid2.columns("descuento")) 'subtotal
   sdx = Val("" & DBGrid2.columns("subtotal")) * Val("" & DBGrid2.columns("igv")) / 100
   DBGrid2.columns("impuesto") = sdx 'igv
   DBGrid2.columns("total") = Val("" & DBGrid2.columns("subtotal")) + sdx 'neto
   sdx = Val("" & DBGrid2.columns("total")) / Val(DBGrid2.columns("cantidad"))
   DBGrid2.columns("precio") = sdx
End If
If racu = "E" Or racu = "N" Then  'si es nota credito por compras o ventas
   If Val("" & DBGrid2.columns("precio")) > 0 Then
      DBGrid2.columns("precio") = -Val("" & DBGrid2.columns("precio"))
   End If
End If
DBGrid2.columns("tpercepcio") = Val(Format(Val("" & DBGrid2.columns("total")) * Val("" & DBGrid2.columns("percepcion")) / 100, "0.00"))
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
DBGrid2.Enabled = False
Command1_Click
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
Command1_Click
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
DBGrid2.Row = fila    'El índice de la primera fila empieza en 0.
suma = suma + Val("" & DBGrid2.columns("descripcio").Value)
Next

End Function
Sub borrar_detalle_todo_registro()
On Error GoTo cmd45_err
ir_primero
amk12:
Data2.Recordset.Delete
Data2.Refresh
GoTo amk12
Exit Sub
cmd45_err:
Exit Sub
End Sub

Sub borrar_detalle_linea()
Data2.Recordset.Delete
DBGrid2.Refresh
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
On Error GoTo cmd17_err
Data1.Recordset.Close
Exit Sub
cmd17_err:
Exit Sub
End Sub
Sub sumar_detalle2()
On Error GoTo cmd34_err
Dim fila As Integer
Dim xtotal As Double
Dim xdescuento As Double
Dim xneto As Double
Dim ximpuesto As Double
Dim xsubtotal As Double
Dim xc1 As Double
Dim xc2 As Double
Dim xc3 As Double
Dim xc4 As Double
Dim xgravado As Double
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
DBGrid2.Row = fila
If "" & DBGrid2.columns(34).Value = "1" Then
   xc1 = xc1 + Val("" & DBGrid2.columns("total").Value)
End If
If "" & DBGrid2.columns(35).Value = "1" Then
   xc2 = xc2 + Val("" & DBGrid2.columns("total").Value)
End If
If "" & DBGrid2.columns(36).Value = "1" Then
   xc3 = xc3 + Val("" & DBGrid2.columns("total").Value)
End If
If "" & DBGrid2.columns(37).Value = "1" Then
   xc4 = xc4 + Val("" & DBGrid2.columns("total").Value)
End If
xntcant = xntcant + Val("" & DBGrid2.columns("cantidad").Value) 'suma bruto
xneto = xneto + Val("" & DBGrid2.columns("neto").Value) 'suma bruto
xdescuento = xdescuento + Val("" & DBGrid2.columns("descuento").Value) 'suma descuento
xsubtotal = xsubtotal + Val("" & DBGrid2.columns("subtotal").Value) ' suma subtotal
ximpuesto = ximpuesto + Val("" & DBGrid2.columns("impuesto").Value) 'suma impuesto
xtotal = xtotal + Val("" & DBGrid2.columns("total").Value)  'suma total
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
Dim fila As Integer
Dim xtotal As Double
Dim xdescuento As Double
Dim xneto As Double
Dim ximpuesto As Double
Dim xsubtotal As Double
Dim xflete As Double
Dim sdx As Double
Dim xc1 As Double
Dim xc2 As Double
Dim xc3 As Double
Dim xc4 As Double
Dim xc5 As Double
Dim xc6 As Double
Dim xc7 As Double
Dim xc8 As Double
Dim xc9 As Double
Dim xpercep As Double

Dim xgravado As Double
Dim vr
Dim xntcant As Double
xpercep = 0
xgravado = 0
xntcant = 0
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
Data2.Recordset.MoveFirst
Do
If Data2.Recordset.EOF Then Exit Do
If Val("" & Data2.Recordset.Fields("igv")) = 0 Then
xgravado = xgravado + Val("" & Data2.Recordset.Fields("total"))
End If
xflete = xflete + Val("" & Data2.Recordset.Fields("cantidad")) * Val("" & Data2.Recordset.Fields("flete")) 'flete
xntcant = xntcant + Val("" & DBGrid2.columns(3).Value) 'suma bruto
xtotal = xtotal + Val("" & Data2.Recordset.Fields("total"))
xdescuento = xdescuento + Val("" & Data2.Recordset.Fields("descuento"))
xneto = xneto + Val("" & Data2.Recordset.Fields("neto"))
ximpuesto = ximpuesto + Val("" & Data2.Recordset.Fields("impuesto"))
xsubtotal = xsubtotal + Val("" & Data2.Recordset.Fields("subtotal"))
xpercep = xpercep + Val("" & Data2.Recordset.Fields("tpercepcio"))
Data2.Recordset.MoveNext
Loop
tflete = Format(xflete, "0.00")
gravado = Format(xgravado, "0.00")
ntcant = Format(xntcant, "0.00")
txtotal = Format(xtotal, "0.00")
txdescuento = Format(xdescuento, "0.00")
txneto = Format(xneto, "0.00")
tximpuesto = Format(ximpuesto, "0.00")
txsubtotal = Format(xsubtotal, "0.00")
txpercepcio = Format(xpercep, "0.00")
sdx = Val(txtotal) + Val(txpercepcio)
xtotper = Format(sdx, "0.00")
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
Image1.Enabled = xsw
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
localf.Enabled = xsw
bodegaf.Enabled = xsw
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
DBGrid2.Enabled = xsw

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
numero.Enabled = xsw

End Sub
Function cargar_registrod()
Dim i As Integer
Dim mytablex As New adodb.Recordset
   mytablex.Open "SELECT * FROM " & dgusuariog & " where  local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & numero & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount = 0 Then 'si existe
      mytablex.Close
      Exit Function
   End If

   Do
   If mytablex.EOF Then Exit Do
         Data2.Recordset.AddNew
         For i = 0 To mytablex.Fields.count - 1
              Data2.Recordset.Fields(i) = mytablex.Fields(i)
         Next i
         Data2.Recordset.Fields("local") = "" & local1
         Data2.Recordset.Fields("tipo") = "" & ttipo
         Data2.Recordset.Fields("serie") = "" & serie
         Data2.Recordset.Fields("numero") = "" & numero
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
Sub proceso_impresion1()
Dim found As Integer
Dim archivot As String
On Error GoTo cmd6_err:
    cerrar_archivo
    factura_formato "" & local1, "" & ttipo, "" & serie, "" & numero, ""
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    Exit Sub
cmd6_err:
    MsgBox "Mensaje, Error al iniciar Impresion " & error$
    Exit Sub
End Sub
Function verifica_doble(buf As String)
Dim mytabley As Table
verifica_doble = 0
Exit Function
Set mytabley = mydbxglo.OpenTable(dgusuario)
mytabley.Index = "cuerpo"
mytabley.Seek "=", ttipo, serie, numero, buf
If Not mytabley.NoMatch Then
   verifica_doble = 1 'estab esto
   'verifica_doble = 0
End If
mytabley.Close
End Function
Sub grabar_cuentaxc()
Dim mytabley As New adodb.Recordset
Dim buf As String
On Error GoTo cmd2340_err



'---------- validando si es cuenta corriente
If valida_flag("" & racu) = 2 Then    'compras
   buf = "cuentap"
   
End If
If valida_flag("" & racu) = 1 Then
   buf = "cuentac"
   
End If
   mytabley.Open "SELECT * FROM " & buf & " where  local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & numero & "'", cn, adOpenKeyset, adLockOptimistic
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
Sub grabar_registro_cuentac(mytabley As adodb.Recordset)
Dim wfecha As String
   mytabley.Fields("fpago") = busca_fpagoc("" & fpago)
   mytabley.Fields("zona") = "" & zona
   mytabley.Fields("grupo") = "C"
   mytabley.Fields("acu") = "" & acu
   mytabley.Fields("local") = "" & local1
   mytabley.Fields("tipo") = "" & ttipo
   mytabley.Fields("serie") = "" & serie
   mytabley.Fields("nombre") = Mid$("" & Label17, 1, 35)
   mytabley.Fields("vendedor") = "" & vendedor
   mytabley.Fields("numero") = "" & numero
   mytabley.Fields("tipoclie") = "" & tipoclie
   mytabley.Fields("codigo") = "" & codigo
   mytabley.Fields("cuota") = "1"
   mytabley.Fields("fecha") = Format("" & fecha, "dd/mm/yyyy")
   mytabley.Fields("moneda") = "" & moneda
   mytabley.Fields("total") = Val("" & txtotal)
   mytabley.Fields("abono") = 0
   mytabley.Fields("saldo") = Val("" & txtotal)
   mytabley.Fields("estado") = "0"
   
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
Dim mytablex As New adodb.Recordset
   mytablex.Open "SELECT * FROM fpago where  fpago='" & buf & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount > 0 Then  'si existe
      busca_fpagoc = "" & mytablex.Fields("tipo")
   End If
mytablex.Close
End Function
Function graba_fpagov()

Dim mytabley As New adodb.Recordset
Dim mytablex As New adodb.Recordset
Dim xyfpago As String
'---------- validando si es cuenta corriente
xyfpago = ""

mytablex.Open "SELECT * FROM fpago where  fpago='" & fpago & "'", cn, adOpenKeyset, adLockOptimistic
If mytablex.RecordCount > 0 Then  'si existe
   xyfpago = "" & mytablex.Fields("tipo")
End If
mytabley.Open "SELECT * FROM fpagov where  local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & numero & "'", cn, adOpenKeyset, adLockOptimistic
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
Sub grabar_registro_fpagov(mytabley As adodb.Recordset)
   mytabley.Fields("local") = "" & local1
   mytabley.Fields("tipo") = "" & ttipo
   mytabley.Fields("serie") = "" & serie
   mytabley.Fields("numero") = "" & numero
   mytabley.Fields("tipoclie") = "" & tipoclie
   mytabley.Fields("codigo") = "" & codigo
   mytabley.Fields("fecha") = Format("" & fecha, "dd/mm/yyyy")
   mytabley.Fields("moneda") = "" & moneda
   mytabley.Fields("total") = Val("" & txtotal)
   mytabley.Fields("recibe") = Val("" & txtotal)
   mytabley.Fields("usuario") = "" & gusuario
   If Len(cajero) > 0 Then
   mytabley.Fields("usuario") = "" & cajero
   End If

   mytabley.Fields("fpago") = "" & fpago
   mytabley.Fields("acu") = "" & racu
   mytabley.Fields("local") = local1 'globalocal
   mytabley.Fields("estado") = "2"
   mytabley.Fields("caja") = caja
   If Len(caja) = 0 Then
   mytabley.Fields("caja") = "00"
End If
   mytabley.Fields("servicio") = servicio
   mytabley.Fields("turno") = turno
   mytabley.Fields("vendedor") = vendedor
End Sub
Sub generar_traslados()


End Sub
Function busca_linea(buf As String)

Dim mytablex As New adodb.Recordset
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
Sub cargar_cotizaciones(xlocal1 As String, xtipo1 As String, xserie1 As String, xnumero1 As String)
Dim mytablex As New adodb.Recordset
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
Sub graba_archivo_detalle(mytablex As adodb.Recordset)
Dim i As Integer
Data2.Recordset.AddNew
For i = 0 To mytablex.Fields.count - 1
    Data2.Recordset.Fields(i) = mytablex.Fields(i)
   Next i

         Data2.Recordset.Fields("tipo") = "" & ttipo
         Data2.Recordset.Fields("serie") = "" & serie
         Data2.Recordset.Fields("numero") = "" & numero
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
Dim mytablex As New adodb.Recordset
   mytablex.Open "SELECT * FROM tipo where   tipo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount > 0 Then  'si existe
   busca_tipo_carga = 1
   Select Case "" & mytablex.Fields("tipodoc")
          Case "A", "B", "C", "D", "G", "E", "F"  'VENTAS
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
Dim rconsulta As New adodb.Recordset
Dim found As Integer
Dim buf As String
found = busca_tipo_carga("" & dbGrid1.columns(0))
If found = 0 Then Exit Sub
buf = "select Producto,Descripcio,Unidad,Factor,Cantidad,Precio,Total,Moneda from " & xarchivo1 & " where tipo='" & dbGrid1.columns(0) & "' and serie='" & dbGrid1.columns(1) & "' and numero='" & dbGrid1.columns(2) & "'"
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
dbGrid1.SetFocus
End Sub
Sub pone_tallas()
t1 = "" & DBGrid2.columns(18)
t2 = "" & DBGrid2.columns(19)
t3 = "" & DBGrid2.columns(20)
t4 = "" & DBGrid2.columns(21)
t5 = "" & DBGrid2.columns(22)
t6 = "" & DBGrid2.columns(23)
t7 = "" & DBGrid2.columns(24)
t8 = "" & DBGrid2.columns(25)
t9 = "" & DBGrid2.columns(26)
t10 = "" & DBGrid2.columns(27)
t11 = "" & DBGrid2.columns(28)
t12 = "" & DBGrid2.columns(29)
t13 = "" & DBGrid2.columns(30)
t14 = "" & DBGrid2.columns(31)
t15 = "" & DBGrid2.columns(32)
t16 = "" & DBGrid2.columns(33)
End Sub
Sub decarga_saldo_talla(mytablex As adodb.Recordset, mytabley As adodb.Recordset, signo As Double)
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
observa1 = "" & DBGrid2.columns("observa1")
observa2 = "" & DBGrid2.columns("observa2")
observa3 = "" & DBGrid2.columns("observa3")
observa4 = "" & DBGrid2.columns("observa4")
End Sub
Sub ingreso_locales()
xxpone_locales
Frame3.Visible = True
observa1.SetFocus
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
Dim sdx As Double
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
Dim i As Integer

Dim mytablex As New adodb.Recordset
Dim mytabley As New adodb.Recordset
Dim sw As Integer
Dim xbodega As String
Dim xsaldo As Double
Dim xbuf As String
Dim xcosto As Double
Dim xcostou As Double
Dim xfactor As Double
Dim xunidad As String
Dim xmargen As Double
On Error GoTo cmd89012_err
For i = 0 To 9
    campo_precios(i).unidad = ""
    campo_precios(i).factor = ""
    campo_precios(i).precio = ""
    campo_precios(i).costo = ""
    campo_precios(i).margen = ""
    campo_precios(i).stock = ""
Next i

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
      xcostou = Val("" & mytablex.Fields("costou")) * Val("" & mytablex.Fields("factor"))
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

DBGrid4.Refresh
Frame5.Visible = True
DBGrid4.SetFocus
Exit Sub
cmd89012_err:
MsgBox "Error en carga Grid " + error$, 48, "Aviso"
Exit Sub

End Sub
Function busca_tipox(buf As String) As String

Dim mytablex As New adodb.Recordset
Dim sdx As Double
'Label16 = ""

   mytablex.Open "SELECT * FROM tipo where   tipo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount > 0 Then  'si existe
      busca_tipox = "" & mytablex.Fields("tipodoc")
   End If
mytablex.Close

End Function
Function valida_flag(buf As String)
Dim mytablex As New adodb.Recordset

   mytablex.Open "SELECT * FROM tipo where  tipo='" & ttipo & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount > 0 Then  'si existe
       Select Case "" & mytablex.Fields("tipodoc")
       Case "Z"
          valida_flag = 3
       Case "T", "A", "B", "C", "D", "G", "E", "F" 'VENTAS
       valida_flag = 1
       Case "S", "J", "K", "L", "M", "P", "N", "O" 'COMPRAS
       valida_flag = 2
       End Select
   End If
mytablex.Close
End Function
Function graba_adelantos(buf1 As String, buf2 As String, buf3 As String, buf4 As String, xsw As String)
Dim mytablex As New adodb.Recordset
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
Sub descarga_el_uso(buf0 As String, buf1 As String, buf2 As String, buf3 As String, xsw As String)
On Error GoTo cmd8912d
Dim mytablex As New adodb.Recordset
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
Function crea_nuevos_proveedores(buf1 As String, buf2 As String, buf3 As String, buf4 As String)
Dim mytablex As New adodb.Recordset
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
Function crea_nuevos_clientes(buf1 As String, buf2 As String, buf3 As String, buf4 As String)
Dim mytablex As New adodb.Recordset

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
Dim mytablex As New adodb.Recordset
If Len(codigo) = 0 Then Exit Function
mytablex.Open "SELECT * FROM codprov where  codigo='" & codigo & "' and codigoP='" & buf2 & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount > 0 Then  'si existe
      buf2 = "" & mytablex.Fields("producto")
      busca_cod_prov = 1
   End If
mytablex.Close

End Function
Function busca_equiva(buf As String) As Integer
Dim mytablex As New adodb.Recordset
If mytablex.State = 1 Then mytablex.Close
   mytablex.Open "SELECT * FROM productb where barras='" & buf & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount > 0 Then
      buf = "" & mytablex.Fields("producto")
      busca_equiva = 1
      mytablex.Close
      Exit Function
   End If
   mytablex.Close
   
   mytablex.Open "SELECT * FROM producto where barras='" & buf & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount > 0 Then
      buf = "" & mytablex.Fields("producto")
      busca_equiva = 1
   End If
   mytablex.Close
End Function

Function busca_caja()
Dim mytablex As New adodb.Recordset

   mytablex.Open "SELECT * FROM parameca where  caja='" & caja & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount > 0 Then  'si existe
      busca_caja = 1
   End If
mytablex.Close

End Function
Function busca_turno()
Dim mytablex As New adodb.Recordset
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
If Len(codigo) = 0 Or Len(serie) = 0 Or Len(numero) = 0 Then ' si es datos principales sin datos solo salir
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
Dim rs As Recordset
Dim i As Integer
Dim pracu As String
Dim buf1 As String
Dim found As Integer
Dim mytablex As New adodb.Recordset
Dim mytablexy As New adodb.Recordset
Dim te As String
Dim ts As String
Dim xc1 As Double
Dim xc2 As Double
Dim xc3 As Double
Dim xc4 As Double
Dim fila As Integer
Dim sw As Integer

sw = 0
'Set mytablexy = mydbxglo.OpenTable(dgusuariog)
'mytablexy.Index = "tdetalle"

      mytablex.Open "SELECT * FROM " & cgusuario & " where  local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & numero & "'", cn, adOpenKeyset, adLockOptimistic
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


      cn.Execute ("delete from " & dgusuariog & " where local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & numero & "'")
      mytablexy.Open "SELECT * FROM " & dgusuariog & " where local='" & local1 & "' and tipo='" & ttipo & "' and serie='" & serie & "' and numero='" & numero & "'", cn, adOpenKeyset, adLockOptimistic
      
Data2.Refresh
Set rs = Data2.Recordset.Clone
Do
If rs.EOF Then Exit Do
mytablexy.AddNew
For i = 0 To rs.Fields.count - 1
mytablexy.Fields(i) = rs.Fields(i)
Next i
mytablexy.Fields("local") = "" & local1
mytablexy.Fields("tipo") = "" & ttipo
mytablexy.Fields("serie") = "" & serie
mytablexy.Fields("numero") = "" & numero
mytablexy.Fields("vendedor") = "" & vendedor
mytablexy.Fields("moneda") = "" & moneda
mytablexy.Fields("bodega") = "" & bodega
mytablexy.Fields("codigo") = "" & codigo
mytablexy.Fields("localf") = "" & localf
mytablexy.Fields("bodegaf") = "" & bodegaf
mytablexy.Fields("acu") = "" & racu
mytablexy.Fields("acu1") = "" & acu1
mytablexy.Fields("flage") = "" & flage
mytablexy.Fields("tipoclie") = tipoclie
mytablexy.Fields("usuario") = "" & gusuario
If Len(cajero) > 0 Then
   mytablexy.Fields("usuario") = "" & cajero
End If

mytablexy.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
mytablexy.Fields("hora") = Format(Now, "hh:MM")
mytablexy.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
mytablexy.Fields("estado") = "0"
mytablexy.Fields("caja") = caja
If Len(caja) = 0 Then
   mytablexy.Fields("caja") = "00"
End If
mytablexy.Fields("turno") = turno
mytablexy.Fields("servicio") = servicio
mytablexy.Update
grabar1 = 1
rs.MoveNext
Loop
mytablexy.Close
End Function
Function ver_cambio_precios(buf As String)
Dim sw As Integer
Dim mytablex As New adodb.Recordset
sw = 0
   mytablex.Open "SELECT * FROM producto where  producto='" & buf & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount > 0 Then  'si existe
   sw = 1
changepr.codigo = "" & mytablex.Fields("producto")
changepr.descripcio = Mid$("" & mytablex.Fields("descripcio"), 1, 60)
changepr.fotonombre = "" & mytablex.Fields("fotonombre")
changepr.monedac = "" & mytablex.Fields("monedac")
changepr.unidad = "" & mytablex.Fields("unidad")
changepr.factor = "" & mytablex.Fields("factor")
changepr.costop = "" & mytablex.Fields("costop")
'changepr.costou = "" & mytablex.Fields("costou")
changepr.costou = "" & DBGrid2.columns(5)
'changepr.ccosto = "" & mytablex.Fields("ccosto")
changepr.monedav = "" & mytablex.Fields("monedav")
End If
mytablex.Close
If sw = 1 Then
   changepr.Show 1
   ver_cambio_precios = 1
End If

End Function
Function ver_docena1(mytablex As Table)
Dim xbuf1(10) As String
Dim xbuf2(10) As Double
Dim xbuf3(10) As Double
Dim j As Integer
Dim i As Integer
Dim sdx As Double

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
      For i = 0 To 9
          If i = 0 Then
             sdx = xbuf2(i)
             j = i
          End If
          If xbuf2(i) > sdx Then
             sdx = xbuf2(i)
             j = i
          End If
      Next i
      If sdx > 1 Then
         DBGrid2.columns("unidad") = xbuf1(j)
         DBGrid2.columns("factor") = xbuf2(j)
         DBGrid2.columns("precio") = xbuf3(j)
         DBGrid2.columns("total") = xbuf3(j)
         DBGrid2.columns("subtotal") = xbuf3(j)

      End If
      If sdx = 0 Then  'no pasa nada
      End If

End Function
Function busca_cajero()
Dim mytablex As New adodb.Recordset
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
Dim mytablex As New adodb.Recordset
mytablex.Open "SELECT * FROM tipo where  tipo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount > 0 Then  'si existe
      Select Case "" & mytablex.Fields("tipodoc")
             Case "A", "B", "C", "D", "G", "E", "F", "J", "K", "L", "M", "P", "N", "O", "S", "T"
                  ve_descarga = 1
      End Select
End If
mytablex.Close

End Function
Sub ver_presenta()
Dim buf As String
Dim mytablex As New adodb.Recordset
buf = "" & DBGrid2.columns("producto")
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
Dim mytablex As New adodb.Recordset
mytablex.Open "SELECT * FROM tipo where  tipo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount > 0 Then  'si existe
      tipo_costo = "" & mytablex.Fields("tipocosto")
End If
mytablex.Close
End Function
Sub actualizar_precios(mytablex As adodb.Recordset)
Dim sw As Integer
On Error GoTo cmd89121_err
Dim mytableyy As New adodb.Recordset
Dim mytabley As New adodb.Recordset
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
Sub calcula_margenes(mytablex As adodb.Recordset, mytabley As adodb.Recordset)
Dim sdx As Double
Dim sdx1 As Double
Dim sdx2 As Double
Dim acostou As String
On Error GoTo cmd786_err
      
sdx = Val("" & mytabley.Fields("costou")) + Val("" & mytabley.Fields("flete"))
acostou = "" & sdx

          If mytabley.Fields("monedac") = "S" Then
             If mytabley.Fields("monedav") = "D" Then
                sdx = Val(acostou) / Val(busca_paridadg(0))
                If sdx <= 0 Then
                   sdx = 1
                End If
                acostou = "" & sdx
             End If
          End If
          If mytabley.Fields("monedac") = "D" Then
             If mytabley.Fields("monedav") = "S" Then
                sdx = Val(acostou) * Val(busca_paridadg(0))
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
Sub pone_margenes(mytablex As adodb.Recordset, mytabley As adodb.Recordset)
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
Sub inicializa_precios(mytablex As adodb.Recordset)
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
