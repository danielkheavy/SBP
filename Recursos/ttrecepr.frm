VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ttrecepr 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Tabla de Formulacion"
   ClientHeight    =   10080
   ClientLeft      =   165
   ClientTop       =   -615
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   15240
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consulta"
      Height          =   9975
      Left            =   0
      TabIndex        =   49
      Top             =   30
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
         TabIndex        =   52
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
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Filtrar"
         Height          =   495
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   240
         Width           =   2175
      End
      Begin MSDataGridLib.DataGrid dbgrid13 
         Height          =   8895
         Left            =   120
         TabIndex        =   53
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
   Begin VB.Frame Frame4 
      Caption         =   "Copia Receta"
      Height          =   4575
      Left            =   30
      TabIndex        =   42
      Top             =   45
      Visible         =   0   'False
      Width           =   6255
      Begin VB.TextBox idformulai 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   46
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox idformulaf 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   45
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Copiar"
         Height          =   615
         Left            =   5040
         TabIndex        =   44
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Close"
         Height          =   615
         Left            =   3960
         TabIndex        =   43
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Formula Inicio"
         Height          =   615
         Left            =   120
         TabIndex        =   48
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label15 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Formula Destino"
         Height          =   615
         Left            =   120
         TabIndex        =   47
         Top             =   960
         Width           =   2895
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4920
         Picture         =   "ttrecepr.frx":0000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frame2"
      Height          =   8175
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   14535
      Begin VB.TextBox observa 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   2280
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   3480
         Width           =   5895
      End
      Begin VB.TextBox Grupo 
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
         TabIndex        =   28
         Text            =   "TF"
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox activo 
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
         MaxLength       =   1
         TabIndex        =   27
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox fecha 
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
         MaxLength       =   10
         TabIndex        =   26
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox id 
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
         TabIndex        =   25
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox descripcio 
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
         TabIndex        =   24
         Top             =   960
         Width           =   6015
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   10560
         Picture         =   "ttrecepr.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Imprimir todo"
         Top             =   1560
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   10560
         Picture         =   "ttrecepr.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   480
         Width           =   1470
      End
      Begin VB.TextBox subgrupo 
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
         TabIndex        =   21
         Top             =   3120
         Width           =   1935
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
         MaxLength       =   15
         TabIndex        =   20
         Top             =   600
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
         MaxLength       =   10
         TabIndex        =   19
         Top             =   1320
         Width           =   975
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
         TabIndex        =   18
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox costo 
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
         Left            =   6120
         MaxLength       =   6
         TabIndex        =   17
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observa"
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
         TabIndex        =   41
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Grupo"
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
         TabIndex        =   40
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Activo"
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
         TabIndex        =   39
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha"
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
         TabIndex        =   38
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Id"
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
         TabIndex        =   37
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   36
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SubGrupo"
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
         TabIndex        =   35
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   34
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   33
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   32
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Formula para N"
         Height          =   375
         Left            =   3240
         TabIndex        =   31
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Image Image4 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4200
         Picture         =   "ttrecepr.frx":149E
         Stretch         =   -1  'True
         Top             =   600
         Width           =   375
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4200
         Picture         =   "ttrecepr.frx":17A8
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RefrescaCostos"
         Height          =   495
         Left            =   6120
         TabIndex        =   30
         Top             =   1680
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Copiar Insumos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12600
      TabIndex        =   15
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Procesos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12600
      TabIndex        =   12
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Componentes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12600
      TabIndex        =   11
      Top             =   1200
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFC0&
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
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
         Picture         =   "ttrecepr.frx":1AB2
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
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Picture         =   "ttrecepr.frx":2CC4
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
         Picture         =   "ttrecepr.frx":3ED6
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
         Picture         =   "ttrecepr.frx":50E8
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
         Picture         =   "ttrecepr.frx":62FA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label pproducto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3600
         TabIndex        =   14
         Top             =   120
         Width           =   105
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consulta"
      Height          =   9495
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   12495
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   9015
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   15901
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
         ColumnCount     =   9
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
               LCID            =   13322
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
               LCID            =   13322
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "factor"
            Caption         =   "Fac"
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
         BeginProperty Column05 
            DataField       =   "Id"
            Caption         =   "Id"
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
         BeginProperty Column06 
            DataField       =   "Grupo"
            Caption         =   "Grupo"
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
            DataField       =   "Subgrupo"
            Caption         =   "Subgrupo"
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
            DataField       =   "Activo"
            Caption         =   "Activo"
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
               ColumnWidth     =   5940.284
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label label13 
      BackColor       =   &H00C0C0C0&
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
      Height          =   5415
      Left            =   14640
      TabIndex        =   13
      Top             =   120
      Width           =   4215
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
   Begin VB.Menu djuer1 
      Caption         =   "&Reporte"
      Begin VB.Menu dlkor1 
         Caption         =   "&0.Reporte de Formulas"
      End
      Begin VB.Menu dk9893 
         Caption         =   "&0.Generador Consultas"
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
Attribute VB_Name = "ttrecepr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txformxu As New ADODB.Recordset

Private Sub ajdu1_Click()

    Dim found As Integer

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa

    If Len(Trim(pproducto)) > 0 Then
        producto = pproducto
        found = busca_productox()

    End If

    Frame2.Visible = True
    Frame2.Caption = "Nuevo"
    cmdGuardar.Enabled = True
    habilita 1
    id.Enabled = False
    id = ""
    producto.SetFocus

End Sub

Private Sub bo712_Click()

    Dim buf As String

    On Error GoTo cmd656_err

    buf = "" & txformxu.Fields("id")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If MsgBox("Desea Borra " + "" & txformxu.Fields("id"), 1, "Aviso") <> 1 Then
        Exit Sub

    End If

    txformxu.Delete
    Command1_Click

    Exit Sub
cmd656_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub buffer_DblClick()
    Command1_Click

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    Command1_Click

End Sub

Private Sub cio9234_Click()

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

    found = grabar()

End Sub

Private Sub cmdPrint_Click()

    'djuer1_Click
End Sub

Private Sub cmdSave_Click()
    f8443_Click

End Sub

Private Sub Command2_Click()

    On Error GoTo cmfd5611_err

    tcompone.productof = "" & txformxu.Fields("producto")
    tcompone.idx = "" & txformxu.Fields("id")
    tcompone.idxnombre = "" & txformxu.Fields("descripcio")
    tcompone.Show 1
    Exit Sub
cmfd5611_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command3_Click()

    On Error GoTo cmfd125611_err

    tproceso.idx = "" & txformxu.Fields("id")
    tproceso.idxnombre = "" & txformxu.Fields("descripcio")
    tproceso.Show 1
    Exit Sub
cmfd125611_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command4_Click()
    filtro

End Sub

Private Sub Command5_Click()

    On Error GoTo cmd9089_err

    idformulaf = Trim("" & txformxu.Fields("id"))
    Frame4.Visible = True
    Exit Sub
cmd9089_err:
    MsgBox "Elija un Dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command6_Click()

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    If Len(Trim(idformulai)) = 0 Then
        MsgBox "Seleccione un dato", 48, "Aviso"
        Exit Sub
   
    End If

    If Not IsNumeric(idformulai) Then
        MsgBox "Seleccione un dato", 48, "Aviso"
        Exit Sub

    End If

    If Len(Trim(idformulaf)) = 0 Then
        MsgBox "Seleccione un dato", 48, "Aviso"
        Exit Sub
   
    End If

    If Not IsNumeric(idformulaf) Then
        MsgBox "Seleccione un dato", 48, "Aviso"
        Exit Sub

    End If
   
    mytablex.Open "select * from componente where id=" & Val(idformulai) & "", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        MsgBox "No existe formula Inicial ", 48, "Aviso"
        Exit Sub

    End If

    mytabley.Open "select * from componente where id=" & Val(idformulaf), cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        If MsgBox("Ya contine componentes,Desea Agregar mas... ", 1, "Aviso") <> 1 Then
            mytablex.Close
            mytabley.Close
            Exit Sub

        End If

    End If

    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew
        mytabley.Fields("producto") = Trim("" & mytablex.Fields("producto"))
        mytabley.Fields("descripcio") = Trim("" & mytablex.Fields("descripcio"))
        mytabley.Fields("unidad") = Trim("" & mytablex.Fields("unidad"))
        mytabley.Fields("factor") = Val("" & mytablex.Fields("factor"))
        mytabley.Fields("cantidad") = Val("" & mytablex.Fields("cantidad"))
        mytabley.Fields("porcentajepeso") = Val("" & mytablex.Fields("porcentajepeso"))
        mytabley.Fields("porcentajemerma") = Val("" & mytablex.Fields("porcentajemerma"))
        mytabley.Fields("costo") = Val("" & mytablex.Fields("costo"))
        mytabley.Fields("moneda") = Trim("" & mytablex.Fields("moneda"))
        mytabley.Fields("tipo") = Trim("" & mytablex.Fields("tipo"))
        mytabley.Fields("explosion") = Trim("" & mytablex.Fields("explosion"))
        mytabley.Fields("formula") = Val("" & mytablex.Fields("formula"))
        mytabley.Fields("productof") = Trim("" & txformxu.Fields("producto"))
        mytabley.Fields("id") = Val("" & idformulaf)
        mytabley.Update
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close
    MsgBox "Proceso Realizado ", 48, "Aviso"
    Frame4.Visible = False
    'proceso de copiar

End Sub

Private Sub Command7_Click()
    Frame4.Visible = False

End Sub

Private Sub dbgrid13_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim found As Integer

    If KeyCode = 27 Then
        Text1.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "1" Then
            Grupo = Trim("" & dbgrid13.columns(1))
            Frame3.Visible = False
            Grupo.SetFocus

        End If

        If opcion1 = "2" Then
            subgrupo = Trim("" & dbgrid13.columns(1))
            Frame3.Visible = False
            subgrupo.SetFocus

        End If

        If opcion1 = "3A" Then
            idformulai = Trim("" & dbgrid13.columns(1))
            Frame3.Visible = False
            idformulai.SetFocus

        End If

        If opcion1 = "3" Then
            producto = Trim("" & dbgrid13.columns(1))
            descripcio = Trim("" & dbgrid13.columns(0))
            unidad = Trim("" & dbgrid13.columns(2))
            factor = Trim("" & dbgrid13.columns(3))
   
            Frame3.Visible = False
            subgrupo.SetFocus

        End If

    End If

End Sub

Private Sub dk9893_Click()

    If Frame2.Visible = True Then Exit Sub
    reporgen.NAMETABLA = "formulacion"
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

    If opcion1 = "1" Then  'bodega
        If Len(Trim(buffer)) = 0 Then
            cad = "SELECT * from formulacion  order by descripcio  "

            If Len(Trim(pproducto)) > 0 Then
                cad = cad & " where producto='" & pproducto & "'"

            End If

        End If
   
        If Len(Trim(buffer)) > 0 Then
            cad = "SELECT *  from formulacion  where "

            If Len(Trim(pproducto)) > 0 Then
                cad = cad & "  producto='" & pproducto & "'"
                cad = cad & " and " & Combo1 & " like '" & buffer & "%'"
            Else
                cad = cad & " " & Combo1 & " like '" & buffer & "%' order by descripcio"

            End If

        End If

        If txformxu.State = 1 Then txformxu.Close
        txformxu.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbGrid1.DataSource = txformxu
        dbGrid1.columns(0).Width = 4000
        dbGrid1.columns(1).Width = 2000

        If txformxu.RecordCount > 0 Then
            dbGrid1.SetFocus

        End If

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

Private Sub dlkor1_Click()
    'reporte
    reporte_excell txformxu

End Sub

Private Sub dlo132_Click()

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

    If Frame4.Visible = True Then
        Frame4.Visible = False
        Exit Sub

    End If

    ttrecepr.Hide
    Unload ttrecepr

End Sub

Private Sub f8443_Click()

    Dim buf As String

    On Error GoTo cmd456_err

    buf = "" & txformxu.Fields("id")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Modifica"
    cmdGuardar.Enabled = True
    pone_registro
    habilita 1
    id.Enabled = False
    descripcio.SetFocus
    Exit Sub
cmd456_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fjh433_Click()

    Dim buf As String

    On Error GoTo cmd556_err

    buf = "" & txformxu.Fields("id")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Zoom"
    cmdGuardar.Enabled = False
    pone_registro
    habilita 1
    id.Enabled = False
    'MsgBox "ABC"
    descripcio.SetFocus
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

    Label13 = Label13 & "Es el recetario o listado de los materiales de los productos A producir" & Chr$(10) & Chr$(13)
    Label13 = Label13 & "Listado de componentes y materiales(Materia prima,insumos,suministros)" & Chr$(10) & Chr$(13)
    Label13 = Label13 & "Indicar la ruta de produccion y el tiempo estimado de produccion" & Chr$(10) & Chr$(13)
    Label13 = Label13 & "Al terminar relacionar con la tabla de productos" & Chr$(10) & Chr$(13)

End Sub

Sub inicializa()
    fecha = Format(Now, "dd/mm/yyyy")
    Grupo = "TF"
    subgrupo = ""
    activo = "S"
    observa = ""
    descripcio = ""
    producto = ""
    unidad = ""
    factor = ""
    costo = ""

End Sub

Sub pone_registro()
    costo = Trim("" & txformxu.Fields("costo"))
    unidad = Trim("" & txformxu.Fields("unidad"))
    factor = Trim("" & txformxu.Fields("factor"))
    fecha = Trim("" & txformxu.Fields("fecha"))
    activo = Trim("" & txformxu.Fields("activo"))
    producto = Trim("" & txformxu.Fields("producto"))
    subgrupo = Trim("" & txformxu.Fields("subgrupo"))
    observa = Trim("" & txformxu.Fields("observa"))
    id = Trim("" & txformxu.Fields("id"))
    descripcio = Trim("" & txformxu.Fields("descripcio"))

End Sub

Sub grabando()
    txformxu.Fields("costo") = Val(costo)
    txformxu.Fields("unidad") = Trim(unidad)
    txformxu.Fields("factor") = Val(factor)
    txformxu.Fields("producto") = Trim(producto)
    txformxu.Fields("fecha") = Trim(fecha)
    txformxu.Fields("activo") = Trim(activo)
    txformxu.Fields("grupo") = Trim(Grupo)
    txformxu.Fields("subgrupo") = Trim(subgrupo)
    txformxu.Fields("observa") = Trim(observa)
    txformxu.Fields("descripcio") = Trim(descripcio)

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
        txformxu.AddNew
        grabando
        txformxu.Update
        guarda_costo
        dlo132_Click
        Exit Function

    End If

    If Frame2.Caption = "Modifica" Then
        grabando
        txformxu.Update
        guarda_costo
        dlo132_Click
        Exit Function

    End If

End Function

Function valida()

    Dim found As Integer

    If Len(Trim(fecha)) <> 10 Then
        fecha.SetFocus
        Exit Function

    End If

    If Not IsDate(fecha) Then
        fecha.SetFocus
        Exit Function

    End If

    If Len(descripcio) = 0 Then
        descripcio.SetFocus
        Exit Function

    End If

    If Len(Grupo) = 0 Then
        Grupo.SetFocus
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

    If Len(subgrupo) = 0 Then
        subgrupo.SetFocus
        Exit Function

    End If

    found = valida_producto()

    If found = 0 Then
        producto.SetFocus
        Exit Function

    End If

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

Private Sub Grupo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_grupo

    End If

End Sub

Private Sub Image1_Click()
    consulta_subgrupo

End Sub

Private Sub Image2_Click()
    consulta_formula

End Sub

Sub consulta_formula()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.AddItem "Producto"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "3A"
    Text1.SetFocus
    'Command4_Click

End Sub

Private Sub Image4_Click()
    consulta_producto

End Sub

Private Sub Label12_Click()
    calcula_costo

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

Sub consulta_grupo()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.AddItem "tablapro"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "1"
    Text1.SetFocus
    Command4_Click

End Sub

Sub consulta_subgrupo()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.AddItem "subtablapro"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "2"
    Text1.SetFocus
    Command4_Click

End Sub

Sub consulta_producto()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.AddItem "Producto"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "3"
    Text1.SetFocus
    'Command4_Click

End Sub

Private Sub producto_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_producto

    End If

End Sub

Private Sub subgrupo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_subgrupo

    End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    Command4_Click

End Sub

Sub filtro()

    Dim mytablex As New ADODB.Recordset

    Dim cad      As String

    If opcion1 = "1" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select Descripcio,Tablapro from Tablapro "

        End If

        If Len(Text1) > 0 Then
            cad = "select Descripcio,Tablapro from tablapro where " & Combo2 & " like '" & Text1.Text & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        dbgrid13.columns(0).Width = 5000
        dbgrid13.columns(1).Width = 2000
        'dbgrid13.columns(2).Width = 1000
        'dbgrid13.columns(3).Width = 1000
           
    End If

    If opcion1 = "3A" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select Descripcio,Id,Producto,Unidad,factor from formulacion order by descripcio "

        End If

        If Len(Text1) > 0 Then
            cad = "select Descripcio,Id,Producto,Unidad,factor from formulacion  where " & Combo2 & " like '" & Text1.Text & "%' order by descripcio"

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
            cad = "select Descripcio,subTablapro,Tablapro from subTablapro where tablapro='" & Grupo & "'"

        End If

        If Len(Text1) > 0 Then
            cad = "select Descripcio,SubTablapro,Tablapro from subtablapro where tablapro='" & Grupo & "' and " & Combo2 & " like '" & Text1.Text & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        dbgrid13.columns(0).Width = 5000
        dbgrid13.columns(1).Width = 2000
        'dbgrid13.columns(2).Width = 1000
        'dbgrid13.columns(3).Width = 1000
           
    End If
   
    If opcion1 = "3" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select Descripcio,Producto,Unidad,factor from Producto "

        End If

        If Len(Text1) > 0 Then
            cad = "select Descripcio,Producto,Unidad,factor from producto where " & Combo2 & " like '" & Text1.Text & "%'"

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

Sub Reporte()

    Dim found As Integer

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    borrar_archivo FileName
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_documento1
    cuerpo_programa_documento1
    '------------------------------------
    Close #1
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Sub cabecera_documento1()

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
    buf = "Reporte de Formulacion  "
    found = formateaa(buf, 90, 2, 0)
    
    found = formateaa("Producto", 8, 0, 0)
    found = formateaa("Descripcio", 60, 0, 0)
    found = formateaa("Und", 7, 0, 0)
    found = formateaa("Factor ", 7, 0, 0)
    found = formateaa("Cant ", 8, 2, 0)
        
    buf = String(150, "-")
    found = formateaa(buf, 150, 2, 0)

End Sub

Sub cuerpo_programa_documento1()

    Dim buf   As String

    Dim found As Integer

    On Error GoTo cmd78812_err

    Do

        If txformxu.EOF Then Exit Do
        buf = "+" & txformxu.Fields("Producto")
        found = formateaa(buf, 7, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txformxu.Fields("descripcio")
        found = formateaa(buf, 59, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txformxu.Fields("unidad")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txformxu.Fields("factor")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 2, 0)
        nlineas
        imprime_receta "" & txformxu.Fields("id")
        txformxu.MoveNext
    Loop
    Exit Sub
cmd78812_err:
    MsgBox "Aviso en cuerpo " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub nlineas()
    contlin = contlin + 1

    If contlin > 45 Then
        cabecera_documento1

    End If

End Sub

Sub imprime_receta(buf1 As String)

    Dim buf      As String

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from componente where id=" & buf1, cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        buf = "-" & mytablex.Fields("Producto")
        found = formateaa(buf, 7, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("descripcio")
        found = formateaa(buf, 59, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("unidad")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("factor")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
    
        buf = "" & mytablex.Fields("cantidad")
        found = formateaa(buf, 7, 0, 0)
        found = formateaa("", 1, 2, 0)
        nlineas
        mytablex.MoveNext
    Loop
    mytablex.Close

End Sub

Function valida_producto()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from producto where producto='" & producto & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        valida_producto = 1

    End If

    mytablex.Close

End Function

Function busca_productox()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from producto where producto='" & producto & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        descripcio = "" & mytablex.Fields("descripcio")
        unidad = "" & mytablex.Fields("unidad")
        factor = "" & mytablex.Fields("factor")

    End If

    mytablex.Close

End Function

Sub reporte_excell(mytablex As ADODB.Recordset)

    Dim found       As Integer

    Dim buf         As String

    Dim buf1        As String

    Dim sdx         As Double

    Dim sdx1        As Double

    Dim Heading(15) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    Dim v           As Long

    Dim h           As Long

    Command1.Visible = True

    On Error GoTo cmd6561245_err
    
    Heading(1) = "Id"
    Heading(2) = "Descripcio"
    Heading(3) = "producto"
    Heading(4) = "Unidad"
    Heading(5) = "factor"
    Heading(6) = "cantidad"
    Heading(7) = "Costo"
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_orden(15, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook

    objExcel.ActiveSheet.Cells(1, 1) = "FECHA HOY  " + Format(Now, "dd/mm/yyyy") & "- HORA HOY  " + Format(Now, "HH:MM:SS")

    v = 4
    h = 1
    sdx1 = 0
    
    Do

        If mytablex.EOF Then Exit Do
        objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("Id")
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("Descripcio")
        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & mytablex.Fields("Producto")
        objExcel.ActiveSheet.Cells(v, h + 3) = "'" & mytablex.Fields("Unidad")
        objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytablex.Fields("factor")
        v = v + 1
        imprime_recetaa mytablex, v, h
        mytablex.MoveNext
    Loop
    Set objExcel = Nothing
    Exit Sub
cmd6561245_err:
    MsgBox "Aviso en reporte orden " + error$, 48, "Aviso"
    Exit Sub

End Sub

Function Formato_orden(Num_Campos As Integer, Nombre_Campos() As String) As Boolean

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
        .columns("A").ColumnWidth = 10
        .columns("B").ColumnWidth = 60
        .columns("C").ColumnWidth = 10
        .columns("D").ColumnWidth = 10
        .columns("E").ColumnWidth = 10
        .columns("F").ColumnWidth = 10
        .columns("G").ColumnWidth = 10
        .columns("H").ColumnWidth = 10
        .columns("i").ColumnWidth = 10
        .columns("j").ColumnWidth = 7
        .columns("k").ColumnWidth = 7
        .columns("l").ColumnWidth = 7

    End With

End Function

Sub imprime_recetaa(mytabley As ADODB.Recordset, v As Long, h As Long)

    Dim mytablex As New ADODB.Recordset

    Dim sw       As Integer

    Dim sdx      As Double

    sw = 0
    sdx = 0
    mytablex.Open "select * from componente where id=" & Val("" & mytabley.Fields("id")), cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        sw = 1
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("Descripcio")
        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & mytablex.Fields("Producto")
        objExcel.ActiveSheet.Cells(v, h + 3) = "'" & mytablex.Fields("Unidad")
        objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytablex.Fields("factor")
        objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytablex.Fields("cantidad")
        objExcel.ActiveSheet.Cells(v, h + 6) = "" & mytablex.Fields("costo")
        sdx = sdx + Val("" & mytablex.Fields("costo")) * Val("" & mytablex.Fields("cantidad"))
        v = v + 1
        mytablex.MoveNext
    Loop
    mytablex.Close

    If sw = 1 Then
        objExcel.ActiveSheet.Cells(v, h + 6) = Format(sdx, "0.00")
        v = v + 1

    End If

End Sub

Sub calcula_costo()

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    If Val("" & id) = 0 Then Exit Sub
    sdx = 0
    mytablex.Open "select * from componente where id=" & Val("" & id), cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + Val("" & mytablex.Fields("costo")) * Val("" & mytablex.Fields("cantidad"))
        mytablex.MoveNext
    Loop
    costo = Format(sdx, "0.00")

End Sub

Sub guarda_costo()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from producto where producto='" & producto & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        mytablex.Fields("costou") = Val(costo) / Val(factor)
        mytablex.Update

    End If

    mytablex.Close

End Sub

