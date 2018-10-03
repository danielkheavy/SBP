VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form xzebra 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SOFTWARE PARA EL ZEBRA "
   ClientHeight    =   9000
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   17550
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "zebra.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   17550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consulta de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   0
      TabIndex        =   29
      Top             =   45
      Visible         =   0   'False
      Width           =   14535
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox BUSCARE 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         MaxLength       =   30
         TabIndex        =   34
         Text            =   "%"
         Top             =   600
         Width           =   4695
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   600
         Width           =   3015
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Filtra"
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
         Left            =   12720
         TabIndex        =   31
         Top             =   360
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid dbgrid2 
         Height          =   6375
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   11245
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
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Copiar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   10080
         TabIndex        =   38
         Top             =   7560
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Familia"
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
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Busqueda"
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
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   3135
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   720
         Left            =   11520
         Picture         =   "zebra.frx":56E82
         Stretch         =   -1  'True
         Top             =   7560
         Width           =   1200
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "BASE DE DATOS IMPRESION"
      Height          =   8415
      Left            =   -30
      TabIndex        =   11
      Top             =   30
      Visible         =   0   'False
      Width           =   14655
      Begin MSDataGridLib.DataGrid dbgrid1 
         Height          =   5895
         Left            =   240
         TabIndex        =   39
         Top             =   1440
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   10398
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   840
         Width           =   6615
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Impresion Normal"
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
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox cantidadd 
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
         Left            =   7200
         MaxLength       =   5
         TabIndex        =   20
         Text            =   "1"
         Top             =   7440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tres Columnas"
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
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Refresca"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   7440
         Width           =   1455
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Busca"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   17
         Top             =   7440
         Width           =   1455
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Borra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   16
         Top             =   7440
         Width           =   1455
      End
      Begin VB.ComboBox text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   7920
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.TextBox tablas 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         MaxLength       =   8
         TabIndex        =   13
         Text            =   "PRODUCTO"
         Top             =   360
         Width           =   6615
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LISTA PRECIOS"
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
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NUMERO COPIAS"
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
         Left            =   5280
         TabIndex        =   22
         Top             =   7440
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "COLUMNAS"
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
         Left            =   120
         TabIndex        =   14
         Top             =   7920
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BASE DE DATOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.PictureBox SSPanel1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   17490
      TabIndex        =   8
      Top             =   0
      Width           =   17550
      Begin VB.Label Label13 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Etiq.Estatico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   26
         Top             =   720
         Width           =   1215
      End
      Begin VB.Image Image8 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   600
         Left            =   4080
         Picture         =   "zebra.frx":5718C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Etiqueta/Bd"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   25
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Diseño"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cargar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   1215
      End
      Begin VB.Image Image7 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   600
         Left            =   120
         Picture         =   "zebra.frx":57496
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label nombre_fichero 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "\ORION.V4\BARRAS\DEMONIO"
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
         Left            =   10080
         TabIndex        =   10
         Top             =   0
         Width           =   3615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "FORMATO SELECCIONADO"
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
         Left            =   7560
         TabIndex        =   9
         Top             =   0
         Width           =   2775
      End
      Begin VB.Image Image5 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   600
         Left            =   5520
         Picture         =   "zebra.frx":57D60
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1200
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   600
         Left            =   2760
         Picture         =   "zebra.frx":5806A
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1200
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   600
         Left            =   1440
         Picture         =   "zebra.frx":58374
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1200
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   8280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Etiqueta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   12615
      Begin VB.TextBox diseno1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   240
         MaxLength       =   4000
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   8895
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cargar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox numcopia 
         Height          =   495
         Left            =   10440
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "1"
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Guardar Como..."
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
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NUMERO DE COPIAS"
         Height          =   495
         Left            =   9240
         TabIndex        =   7
         Top             =   2160
         Width           =   1575
      End
   End
   Begin MSComDlg.CommonDialog CmDialog1 
      Left            =   11040
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "FoxPro 2.6;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Menu ldso232 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "xzebra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim h           As Long

Dim posicionxx  As Single

Dim posicion    As Single

Dim countlin    As Integer

Dim countetiq   As Integer

Dim NumOfCopies As String

Dim Arr()       As String

Dim LastIndex   As Long

Dim Port        As String

Dim ax          As Single

Public Enum eTipoEtiqueta

    cPEQUENA
    Cgrande
    Cmediana
    Cdonas
    cbarras
    Cplu

End Enum

Public Enum eOrientation

    C0
    C90
    c180
    c270

End Enum

Dim dbzebra       As New ADODB.Recordset

Dim mytablexx     As New ADODB.Recordset

Dim mOrientation  As String

Dim SendToPrinter As Boolean

Const DEBUGGING = True

Private Sub BUSCARE_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        Command10_Click

    End If

End Sub

Private Sub Claves_KeyPress(KeyAscii As Integer)

End Sub

Private Sub Combo1_Click()
    Command10_Click

End Sub

Private Sub Command1_Click()

    Dim vr As Integer

    On Error GoTo cmd1123_err

    If Len(Trim(diseno1)) = 0 Then Exit Sub
    If Val(cantidadd) = 0 Then Exit Sub

    '-------------------- IMPRESION DE DISENO DE ETIQUETAS ------------------
    If dbzebra.RecordCount = 0 Then Exit Sub
    dbzebra.MoveFirst
    Do

        If dbzebra.EOF Then Exit Do
        Call SetPort("LPT1")
        Call SetOrientation(C0)
        Call OpenPrinter
        Call BeginLabel
        Call SetNumOfCopies(CSng("" & dbzebra.Fields("acanti")))
        posicion = 0
        Call imprimir_datos
        Call EndLabel
        Call ClosePrinter
        'If ImprimirEtiquetas(Cdonas, "" & dbzebra.Fields("descripcio"), "" & dbzebra.Fields("marca"), "" & dbzebra.Fields("pventa1"), "" & dbzebra.Fields("pventa1"), Val(cantidadd), "57 MTS", "LPT1:", C0) Then
        '     MsgBox "Impresión exitosa"
        'Else
        'MsgBox "Se encontró un error al intentar imprimir"
        'End If
        vr = DoEvents()
        dbzebra.MoveNext
    Loop
    Exit Sub
finish:
    Call EndLabel
    Call ClosePrinter
    '------------------------------------------------------------------------
    
    Exit Sub
cmd1123_err:
    MsgBox "command1 Seleccion un dato ", 48, "Aviso"
    Exit Sub
    
End Sub

Private Sub trescolumnas()

    Dim vr       As Integer

    Dim columnas As Integer

    Dim copias   As Double

    '-------------------- IMPRESION DE DISENO DE ETIQUETAS ------------------
    'ax = 250
    If dbzebra.RecordCount = 0 Then Exit Sub
    dbzebra.MoveFirst
   
    posicion = 0
    copias = 1
    Call SetPort("LPT1")
    Call SetOrientation(C0)
    Call OpenPrinter
    Call BeginLabel
    Call SetNumOfCopies(1)
    Do

        'VERIFICANDO COPIAS
        If dbzebra.EOF Then Exit Do
        Do

            'MsgBox posicion
            'If copias > Val(cantidadd) Then
            If copias > Val("" & dbzebra.Fields("acanti")) Then
                Exit Do

            End If

            If posicion > 460 Then
                posicion = 0
                Call EndLabel
                Call ClosePrinter
                Call SetPort("LPT1")
                Call SetOrientation(C0)
                Call OpenPrinter
                Call BeginLabel
                Call SetNumOfCopies(1)

            End If

            Call imprimir_datos
            copias = copias + 1
            posicion = posicion + 230
        Loop
        copias = 1

        If posicion > 460 Then
            posicion = 0
            Call EndLabel
            Call ClosePrinter
            Call SetPort("LPT1")
            Call SetOrientation(C0)
            Call OpenPrinter
            Call BeginLabel
            Call SetNumOfCopies(1)

        End If

        vr = DoEvents()
        dbzebra.MoveNext
    Loop
    Call EndLabel
    Call ClosePrinter
    Exit Sub
    '------------------------------------------------------------------------
    
    Exit Sub

End Sub

Private Sub SetOrientation(o As eOrientation)

    Dim strO As String

    mOrientation = "N"

    Select Case o

        Case c180
            mOrientation = "I"

        Case c270
            mOrientation = "B"

        Case C90
            mOrientation = "R"

    End Select

End Sub

Private Sub PrintOrientation()

    If mOrientation <> "N" Then
        Call Add("^FW" & mOrientation)

    End If

End Sub

Private Sub BeginLabel(Optional blnSendToPrinter As Boolean = True)
    
    If DEBUGGING Then

        'blnSendToPrinter = False
        'blnSendToPrinter = true
    End If

    SendToPrinter = blnSendToPrinter
    Call Add("^XA", True)

End Sub

Private Sub Add(valor As String, Optional ClearArray As Boolean = False)

    If valor <> "" Then
        If SendToPrinter Then
            Print #h, valor
            LastIndex = -1
        Else

            If ClearArray Then
                ReDim Arr(0) As String
                LastIndex = 0
            Else
                On Local Error Resume Next
                
                LastIndex = UBound(Arr)

                If Err.Number <> 0 Then
                    LastIndex = -1

                End If

                LastIndex = LastIndex + 1
                On Local Error GoTo 0

            End If
            
            ReDim Preserve Arr(LastIndex) As String
            Arr(LastIndex) = valor

        End If

    End If

End Sub

Public Sub CalibrarImpresora(strPuerto As String)
    Call SetPort(strPuerto)
    Call OpenPrinter
    Call BeginLabel
                              
    Call confCalibrate
            
    Call EndLabel
    Call ClosePrinter

End Sub

Private Sub ClosePrinter()

    Dim I As Long

    If LastIndex >= 0 Then

        For I = 1 To 20
            Debug.Print
        Next

        For I = LBound(Arr) To UBound(Arr)
            'Print #h, Arr(i)
            Debug.Print Arr(I)
        Next

    End If

    Close #h

End Sub

Private Sub EndLabel()
    Call Add(NumOfCopies)
    Call Add("^XZ")

End Sub

Private Function FileExist(FileName As String) As Boolean

    Dim X As Long

    On Local Error Resume Next
    X = FreeFile
    Open FileName For Binary As X

    If Err.Number = 0 Then
        FileExist = True
        Close #X
    Else
        FileExist = False

    End If
    
    On Local Error GoTo 0

End Function

Public Function ImprimirEtiquetas(tipo As eTipoEtiqueta, _
                                  codigo As String, _
                                  descripcion As String, _
                                  color As String, _
                                  Lote As String, _
                                  Optional NumCopias As Integer = 1, _
                                  Optional cantidad As String = "100 METROS", _
                                  Optional strPuerto As String = "COM1:", _
                                  Optional lOrientation As eOrientation = C0) As Boolean

    On Local Error GoTo ErrDrv
    
    Call SetPort(strPuerto)
    Call SetOrientation(lOrientation)
    
    Select Case tipo

        Case Cgrande
            Call ImprimirEtiquetaGde(codigo, descripcion, color, Lote, NumCopias, cantidad)

        Case cPEQUENA

            'Call ImprimirEtiquetaPeq(Codigo, Descripcion, Color, Lote, NumCopias, Cantidad)
        Case Cmediana
            Call ImprimirEtiquetaMed(codigo, descripcion, color, Lote, NumCopias, cantidad)

        Case Cdonas
            Call imprimiretiquetadona(codigo, descripcion, color, Lote, NumCopias, cantidad)

        Case cbarras
            Call imprimiretiquetacb(codigo, descripcion, color, Lote, NumCopias, cantidad)

        Case Cplu
            Call imprimiretiquetadona1(codigo, descripcion, color, Lote, NumCopias, cantidad)

    End Select

    ImprimirEtiquetas = True
    Exit Function
    
ErrDrv:
    ImprimirEtiquetas = False

End Function

Private Sub LoadImage(path As String, FileName As String)

    Dim X          As Long

    Dim CurByte    As Byte, I As Long, j As Long

    Dim TotalBytes As Long ', BytesPerRow As Long

    Dim Line       As String, BytesPerRow As Byte

    If FileExist(path & "\" & FileName & ".PCX") Then
        X = FreeFile
        Open path & "\" & FileName & ".PCX" For Binary As X
        
        TotalBytes = LOF(X) - 128       'Cabezera de los PCX: 128 bytes
        Get #X, 67, BytesPerRow         'Byte 67: # Bytes por Linea gráfica
        Call Add("~DGR:" & FileName & ".GRF," & TotalBytes & "," & BytesPerRow & ",")
        Get #X, 128, CurByte
        
        Do
        
            Line = ""

            For j = 1 To BytesPerRow
                CurByte = 0
                Get #X, , CurByte

                If Not EOF(X) Then
                    If Len(Hex(CurByte)) = 1 Then
                        Line = Line & "0" & Hex(CurByte)
                    Else
                        Line = Line & Hex(CurByte)

                    End If

                Else
                    Exit For

                End If

            Next
            Call Add(Line)
        
        Loop While Not EOF(X)
        
        Close #X
        
    End If

End Sub

Private Sub PrintData(Data As String)
    Call PrintOrientation
    Call Add("^FD" & Data & "^FS")

End Sub

Private Sub PrintImage(FileName As String, _
                       Optional xfactor As Integer = 1, _
                       Optional YFactor As Integer = 1)

    Dim Extra As String

    Call PrintOrientation
    Call Add("^XGR:" & FileName & "," & xfactor & "," & YFactor & "^FS")

End Sub

Private Sub SetBlock(Optional Width As String = "0", _
                     Optional MaxLines As String = "1", _
                     Optional AddOrDeleteSpace As String = "0", _
                     Optional Justify As String = "C", _
                     Optional InnerMargin As String = "0")
    Call PrintOrientation
    Call Add("^FB" & Width & "," & MaxLines & "," & AddOrDeleteSpace & "," & Justify & "," & InnerMargin)

    '^FBa,b,c,d,e: Bloque de texto. Precede al ^FD
    '               a=Ancho del texto (0..9999)
    '               b=Número máximo de líneas (1..9999)
    '               c=Agregar o eliminar espacio entre líneas (-9999..9999)
    '               d=justificacion (L,C,R,J)
    '               e=sangría de la segunda línea y sucesivas.  (0..9999)
End Sub

Private Sub OpenPrinter()
    LastIndex = -1
    Call ClosePrinter
    h = FreeFile
    Open Port For Output As h

End Sub

Private Sub PrintBarCode(BarCode As String, _
                         Optional Orientation As eOrientation, _
                         Optional Height As String = "100", _
                         Optional PrintCode As String = "N", _
                         Optional CodeOnTop As String = "N", _
                         Optional PrintCheckDigit As String = "N")

    Dim strO As String

    strO = "N"

    Select Case Orientation

        Case c180
            strO = "R"

        Case c270
            strO = "I"

        Case C90
            strO = "B"

    End Select

    'Call Add("^BA" & strO & "," & Height & "," & PrintCode & "," & CodeOnTop & "," & PrintCheckDigit & "^FD" & BarCode & "^FS")
    Call Add("^BC" & strO & "," & Height & "," & PrintCode & "," & CodeOnTop & "," & PrintCheckDigit & "^FD" & BarCode & "^FS")

    'o = Orientación(N, R, i, B)
    'h=Altura (1..32000)
    'f=Imprimir Línea de Interpretación (Y,N)
    'g=Imprimir Línea de Interpretación sobre el código (Y,N)
    'e=Imprimir dígito de chequeo
End Sub

Private Sub PrintBox(Optional Width As Single = 0, _
                     Optional Height As Single = 0, _
                     Optional BorderThickness As Single = 2)

    Select Case mOrientation

        Case "N", "B"
            Call Add("^GB" & Width & "," & Height & "," & BorderThickness & "^FS")

        Case "R", "I"
            Call Add("^GB" & Height & "," & Width & "," & BorderThickness & "^FS")

    End Select

End Sub

Private Sub SetFont(Optional FontName As String = "A", _
                    Optional Orientation As String = "N", _
                    Optional Height As Single = 14, _
                    Optional Width As Single = 12)
    Call Add("^A" & FontName & Orientation & "," & Height & "," & Width)

End Sub

Private Sub nuevo_diseno()
    Call SetPort("LPT1")
    Call SetOrientation(C0)
    Call OpenPrinter
    Call BeginLabel
    Call SetNumOfCopies(1)

    Call SetPos(10, 10)
    Call SetFont("A", , 14, 12)
    Call SetBlock(700, , , "L")
    Call PrintData("Hola Mundo")
    Call SetPos(50, 10)
    Call PrintBarCode("12345", , 50, "N", "N")

    Call SetPos(10, 250)
    Call SetFont("A", , 14, 12)
    Call SetBlock(700, , , "L")
    Call PrintData("Hola Mundo 1")
    Call SetPos(50, 250)
    Call PrintBarCode("12345", , 50, "N", "N")

    Call SetPos(10, 450)
    Call SetFont("A", , 14, 12)
    Call SetBlock(700, , , "L")
    Call PrintData("Hola Mundo 1")
    Call SetPos(50, 450)
    Call PrintBarCode("12345", , 50, "N", "N")

    Call EndLabel
    Call ClosePrinter

End Sub

Private Sub imprimiretiquetacb(codigo As String, _
                               descripcion As String, _
                               color As String, _
                               Lote As String, _
                               Optional NumCopias As Integer = 1, _
                               Optional cantidad As String = "100 METROS")

    Static BytesPerRow       As Long

    Dim TipoLetraDescripcion As String

    Dim AltoDescripcion      As Single

    Dim AnchoDescripcion     As Single

    Dim XDescripcion         As Single

    Dim YDescripcion         As Single

    Dim TipoLetraColor       As String

    Dim AltoColor            As Single

    Dim AnchoColor           As Single

    Dim XColor               As Single

    Dim YColor               As Single

    Dim TipoLetraCantidad    As String

    Dim AltoCantidad         As Single

    Dim AnchoCantidad        As Single

    Dim XCantidad            As Single

    Dim YCantidad            As Single

    Dim TipoLetraCRESMAR     As String

    Dim AltoCRESMAR          As Single

    Dim AnchoCRESMAR         As Single

    Dim XCRESMAR             As Single

    Dim YCRESMAR             As Single

    Dim TipoLetraLOTE        As String

    Dim AltoLOTE             As Single

    Dim AnchoLOTE            As Single

    Dim XLOTE                As Single

    Dim YLOTE                As Single

    TipoLetraLOTE = "F"
    AltoLOTE = 14
    AnchoLOTE = 12
    YLOTE = 20
    XLOTE = 580
    
    TipoLetraCRESMAR = "F"
    AltoCRESMAR = 14
    AnchoCRESMAR = 12
    YCRESMAR = 20
    XCRESMAR = 30
    
    TipoLetraColor = "F"
    AltoColor = 70
    AnchoColor = 65
    YColor = 200
    XColor = 30
    
    TipoLetraCantidad = "F"
    AltoCantidad = 70
    AnchoCantidad = 30
    YCantidad = 200
    XCantidad = 150
    
    TipoLetraDescripcion = "F"
    AltoDescripcion = 80
    AnchoDescripcion = 70
    YDescripcion = 120
    XDescripcion = 30
    
    Call OpenPrinter
    Call BeginLabel
    Call SetNumOfCopies(NumCopias)
                    
    Call SetPos(10, 10)
    Call SetFont("A", , AltoCRESMAR, AnchoCRESMAR)
    Call SetBlock(600, , , "L")
    Call PrintData(Mid$(codigo, 1, 14))
                    
    Call SetPos(30, 10)
    Call SetFont("A", , AltoCRESMAR, AnchoCRESMAR)
    Call SetBlock(600, , , "L")
    Call PrintData(Mid$(codigo, 15, 6))
    Call SetPos(50, 10)
    Call PrintBarCode(color, , 50, "N", "N")
    Call SetPos(110, 10)  'Relativo al Bar Code
    Call SetFont("A")
    Call PrintData(color)
                    
    '------------
                    
    'Call SetPos(120, 20)
    'Call PrintBarCode(Descripcion, C90, 80, "N", "N")
                    
    '                'Imprimiendo Código de Barra y Logo
    '                Call SetPos(660, 30)
    '                    Call SetFont("F", , 120, 20)
    '                    Call SetBlock(308)
    '                    Call PrintData("DONAS")
    '                    'Call LoadImage(App.Path, "CRESMAR")
    '                    'Call PrintImage("CRESMAR", 2, 2)
            
    Call EndLabel
    Call ClosePrinter

End Sub

Private Sub imprimiretiquetadona1(codigo As String, _
                                  descripcion As String, _
                                  color As String, _
                                  Lote As String, _
                                  Optional NumCopias As Integer = 1, _
                                  Optional cantidad As String = "100 METROS")

    Static BytesPerRow       As Long

    Dim TipoLetraDescripcion As String

    Dim AltoDescripcion      As Single

    Dim AnchoDescripcion     As Single

    Dim XDescripcion         As Single

    Dim YDescripcion         As Single

    Dim TipoLetraColor       As String

    Dim AltoColor            As Single

    Dim AnchoColor           As Single

    Dim XColor               As Single

    Dim YColor               As Single

    Dim TipoLetraCantidad    As String

    Dim AltoCantidad         As Single

    Dim AnchoCantidad        As Single

    Dim XCantidad            As Single

    Dim YCantidad            As Single

    Dim TipoLetraCRESMAR     As String

    Dim AltoCRESMAR          As Single

    Dim AnchoCRESMAR         As Single

    Dim XCRESMAR             As Single

    Dim YCRESMAR             As Single

    Dim TipoLetraLOTE        As String

    Dim AltoLOTE             As Single

    Dim AnchoLOTE            As Single

    Dim XLOTE                As Single

    Dim YLOTE                As Single

    TipoLetraLOTE = "F"
    AltoLOTE = 14
    AnchoLOTE = 12
    YLOTE = 20
    XLOTE = 580
    
    TipoLetraCRESMAR = "F"
    AltoCRESMAR = 14
    AnchoCRESMAR = 12
    YCRESMAR = 20
    XCRESMAR = 30
    
    TipoLetraColor = "F"
    AltoColor = 70
    AnchoColor = 65
    YColor = 200
    XColor = 30
    
    TipoLetraCantidad = "F"
    AltoCantidad = 70
    AnchoCantidad = 30
    YCantidad = 200
    XCantidad = 150
    
    TipoLetraDescripcion = "F"
    AltoDescripcion = 80
    AnchoDescripcion = 70
    YDescripcion = 120
    XDescripcion = 30
    
    Call OpenPrinter
    Call BeginLabel
    Call SetNumOfCopies(NumCopias)
                    
    Call SetPos(20, 20)
    Call SetFont("F,1", , AltoCRESMAR, AnchoCRESMAR)
    Call SetBlock(600, , , "L")
    Call PrintData(Mid$(codigo, 1, 25))
                    
    Call SetPos(50, 20)
    Call SetFont("F,1", , AltoCRESMAR, AnchoCRESMAR)
    Call SetBlock(600, , , "L")
    Call PrintData(Mid$(codigo, 26, 25))
                    
    Call SetPos(100, 20)
    Call SetFont("D,58", , 20, 16)
    Call SetBlock(600, , , "L")
    Call PrintData("S/. " & Format(Val(color), "0.00"))
                    
    Call SetPos(160, 20)
    Call SetFont("D,54", , 18, AnchoCRESMAR)
    Call SetBlock(600, , , "L")
    Call PrintData("PLU:" & descripcion)
                    
    '                'Imprimiendo Código de Barra y Logo
    '                Call SetPos(660, 30)
    '                    Call SetFont("F", , 120, 20)
    '                    Call SetBlock(308)
    '                    Call PrintData("DONAS")
    '                    'Call LoadImage(App.Path, "CRESMAR")
    '                    'Call PrintImage("CRESMAR", 2, 2)
            
    Call EndLabel
    Call ClosePrinter

End Sub

Private Sub imprimiretiquetadona(codigo As String, _
                                 descripcion As String, _
                                 color As String, _
                                 Lote As String, _
                                 Optional NumCopias As Integer = 1, _
                                 Optional cantidad As String = "100 METROS")

    Static BytesPerRow       As Long

    Dim TipoLetraDescripcion As String

    Dim AltoDescripcion      As Single

    Dim AnchoDescripcion     As Single

    Dim XDescripcion         As Single

    Dim YDescripcion         As Single

    Dim TipoLetraColor       As String

    Dim AltoColor            As Single

    Dim AnchoColor           As Single

    Dim XColor               As Single

    Dim YColor               As Single

    Dim TipoLetraCantidad    As String

    Dim AltoCantidad         As Single

    Dim AnchoCantidad        As Single

    Dim XCantidad            As Single

    Dim YCantidad            As Single

    Dim TipoLetraCRESMAR     As String

    Dim AltoCRESMAR          As Single

    Dim AnchoCRESMAR         As Single

    Dim XCRESMAR             As Single

    Dim YCRESMAR             As Single

    Dim TipoLetraLOTE        As String

    Dim AltoLOTE             As Single

    Dim AnchoLOTE            As Single

    Dim XLOTE                As Single

    Dim YLOTE                As Single

    TipoLetraLOTE = "F"
    AltoLOTE = 14
    AnchoLOTE = 12
    YLOTE = 20
    XLOTE = 580
    
    TipoLetraCRESMAR = "F"
    AltoCRESMAR = 14
    AnchoCRESMAR = 12
    YCRESMAR = 20
    XCRESMAR = 30
    
    TipoLetraColor = "F"
    AltoColor = 70
    AnchoColor = 65
    YColor = 200
    XColor = 30
    
    TipoLetraCantidad = "F"
    AltoCantidad = 70
    AnchoCantidad = 30
    YCantidad = 200
    XCantidad = 150
    
    TipoLetraDescripcion = "F"
    AltoDescripcion = 80
    AnchoDescripcion = 70
    YDescripcion = 120
    XDescripcion = 30
    
    Call OpenPrinter
    Call BeginLabel
    Call SetNumOfCopies(NumCopias)
                    
    Call SetPos(20, 20)
    Call SetFont("D", , AltoCRESMAR, AnchoCRESMAR)
    Call SetBlock(600, , , "L")
    Call PrintData(Mid$(codigo, 1, 25))
                    
    Call SetPos(60, 20)
    Call SetFont("D", , AltoCRESMAR, AnchoCRESMAR)
    Call SetBlock(600, , , "L")
    Call PrintData(Mid$(codigo, 26, 25))
                    
    Call SetPos(100, 20)
    Call SetFont("D,54", , 16, AnchoCRESMAR)
    Call SetBlock(600, , , "L")
    Call PrintData(descripcion)
                    
    Call SetPos(160, 20)
    Call SetFont("G", , 10, AnchoCRESMAR)
    Call SetBlock(600, , , "L")
    Call PrintData("S/ " & Format(Val(color), "0.00"))
                    
    '                'Imprimiendo Código de Barra y Logo
    '                Call SetPos(660, 30)
    '                    Call SetFont("F", , 120, 20)
    '                    Call SetBlock(308)
    '                    Call PrintData("DONAS")
    '                    'Call LoadImage(App.Path, "CRESMAR")
    '                    'Call PrintImage("CRESMAR", 2, 2)
            
    Call EndLabel
    Call ClosePrinter

End Sub

Private Sub ImprimirEtiquetaMed(codigo As String, _
                                descripcion As String, _
                                color As String, _
                                Lote As String, _
                                Optional NumCopias As Integer = 1, _
                                Optional cantidad As String = "100 METROS")

    Static BytesPerRow       As Long

    Dim TipoLetraDescripcion As String

    Dim AltoDescripcion      As Single

    Dim AnchoDescripcion     As Single

    Dim XDescripcion         As Single

    Dim YDescripcion         As Single

    Dim TipoLetraColor       As String

    Dim AltoColor            As Single

    Dim AnchoColor           As Single

    Dim XColor               As Single

    Dim YColor               As Single

    Dim TipoLetraCantidad    As String

    Dim AltoCantidad         As Single

    Dim AnchoCantidad        As Single

    Dim XCantidad            As Single

    Dim YCantidad            As Single

    Dim TipoLetraCRESMAR     As String

    Dim AltoCRESMAR          As Single

    Dim AnchoCRESMAR         As Single

    Dim XCRESMAR             As Single

    Dim YCRESMAR             As Single

    Dim TipoLetraLOTE        As String

    Dim AltoLOTE             As Single

    Dim AnchoLOTE            As Single

    Dim XLOTE                As Single

    Dim YLOTE                As Single

    '--------------------------------------------
    '----------------------7.6   3.8-------------
    '--------------------------------------------
    TipoLetraLOTE = "F"
    AltoLOTE = 14
    AnchoLOTE = 12
    YLOTE = 20
    XLOTE = 580
    
    TipoLetraCRESMAR = "A"
    AltoCRESMAR = 14
    AnchoCRESMAR = 12
    YCRESMAR = 20
    XCRESMAR = 30
    
    TipoLetraColor = "O"
    AltoColor = 70
    AnchoColor = 65
    YColor = 200
    XColor = 30
    
    TipoLetraCantidad = "D"
    AltoCantidad = 70
    AnchoCantidad = 30
    YCantidad = 200
    XCantidad = 150
    
    TipoLetraDescripcion = "O"
    AltoDescripcion = 80
    AnchoDescripcion = 70
    YDescripcion = 120
    XDescripcion = 30
    
    Call OpenPrinter
    Call BeginLabel
    Call SetNumOfCopies(NumCopias)
    Call SetPos(40, 40)
                    
    Call SetPos(YCRESMAR, XCRESMAR)
    Call SetFont(TipoLetraCRESMAR, , AltoCRESMAR, AnchoCRESMAR)
    Call SetBlock(600, , , "L")
    Call PrintData(descripcion)
                    
    Call SetPos(YLOTE, XLOTE)
    Call SetFont(TipoLetraLOTE)
    Call SetBlock(200, 1)
    Call PrintData("LOTE/FECHA")
    Call SetPos(YLOTE + 40, XLOTE - 50)
    Call SetFont(TipoLetraLOTE)
    Call SetBlock(300, 1)
    Call PrintData(Lote)
                        
    Call SetPos(YDescripcion, XDescripcion)     ' 190,30 antes
    Call SetFont(TipoLetraDescripcion, , AltoDescripcion, AnchoDescripcion)     '"F", , 34, 34)
    Call SetBlock(770, 2, , "L")
    Call PrintData(descripcion)
                        
    Call SetPos(YColor, XColor)  '190 + 90, 30)
    Call SetFont(TipoLetraColor, , AltoColor, AnchoColor)   '"F", ,34,32 antes
    Call SetBlock(770, 1, , "L")
    Call PrintData(color)
                        
    Call SetPos(YCantidad, XCantidad) '280,250 antes
    Call SetFont(TipoLetraCantidad, , AltoCantidad, AnchoCantidad)    ',"10" antes
    Call SetBlock(770, 1)
    Call PrintData(cantidad)

    Call SetPos(300, 20)
    Call PrintBarCode(codigo, , 80, "N", "N")
    Call SetPos(300 - 20, 20 + 40)  'Relativo al Bar Code
    Call SetFont("B")
    Call PrintData(codigo)
                    
    Call SetFont
    Call SetPos(390, 20)
    Call SetBlock(770, , , "R")
    Call PrintData("HECHO EN PERU")
                    
    Call SetFont
    Call SetPos(390, 20)
    Call SetBlock(770, , , "L")

    If cantidad = "100 MTS" Then
        Call PrintData("UPO-FR005")
    Else
        Call PrintData("UPO-FR006")

    End If

    '                'Imprimiendo Código de Barra y Logo
    '                Call SetPos(660, 30)
    '                    Call SetFont("F", , 120, 20)
    '                    Call SetBlock(308)
    '                    Call PrintData("DONAS")
    '                    'Call LoadImage(App.Path, "CRESMAR")
    '                    'Call PrintImage("CRESMAR", 2, 2)
            
    Call EndLabel
    Call ClosePrinter

End Sub

Private Sub ImprimirEtiquetaGde(codigo As String, _
                                descripcion As String, _
                                color As String, _
                                Lote As String, _
                                Optional NumCopias As Integer = 1, _
                                Optional cantidad As String = "100 METROS")

    Static BytesPerRow       As Long

    Dim TipoLetraDescripcion As String

    Dim AltoDescripcion      As Single

    Dim AnchoDescripcion     As Single

    Dim XDescripcion         As Single

    Dim YDescripcion         As Single

    Dim TipoLetraColor       As String

    Dim AltoColor            As Single

    Dim AnchoColor           As Single

    Dim XColor               As Single

    Dim YColor               As Single

    Dim TipoLetraLOTE        As String

    Dim AltoLOTE             As Single

    Dim AnchoLOTE            As Single

    Dim XLOTE                As Single

    Dim YLOTE                As Single

    Dim TipoLetraCRESMAR     As String

    Dim AltoCRESMAR          As Single

    Dim AnchoCRESMAR         As Single

    Dim XCRESMAR             As Single

    Dim YCRESMAR             As Single
    
    TipoLetraDescripcion = "O"
    AltoDescripcion = 130
    AnchoDescripcion = 90
    YDescripcion = 290
    XDescripcion = 80
    
    TipoLetraCRESMAR = "G"
    AltoCRESMAR = 120
    AnchoCRESMAR = 30
    YCRESMAR = 590
    XCRESMAR = 80
    
    TipoLetraColor = "O"
    AltoColor = 120
    AnchoColor = 85
    YColor = 280
    XColor = 80
    
    TipoLetraLOTE = "F"
    AltoLOTE = 14
    AnchoLOTE = 12
    YLOTE = 700
    XLOTE = 725
    Call OpenPrinter
    Call BeginLabel
    Call SetNumOfCopies(NumCopias)
                
    Call SetPos(40, 40)
    Call PrintBox(975, 720, 4)     'Marco Principal
                    
    Call SetPos(YCRESMAR, XCRESMAR)
    Call SetFont(TipoLetraCRESMAR, , AltoCRESMAR, AnchoCRESMAR)
    Call SetBlock(600, , , "L")
    Call PrintData("CRESMAR")
                    
    Call SetPos(YLOTE, XLOTE)
    Call SetFont(TipoLetraLOTE, , AltoLOTE, AnchoLOTE)
    Call SetBlock(300, 1)
    Call PrintData("LOTE/FECHA")
                        
    Call SetPos(YLOTE - 50, XLOTE)
    Call SetFont(TipoLetraLOTE, , AltoLOTE, AnchoLOTE)
    Call SetBlock(300, 1)
    Call PrintData(Lote)
                                        
    Call SetPos(YDescripcion, XDescripcion)
    Call SetFont(TipoLetraDescripcion, , AltoDescripcion, AnchoDescripcion)
    Call SetBlock(900, 2, , "L")
    Call PrintData(descripcion)
                        
    Call SetPos(YColor, XColor)
    Call SetFont(TipoLetraColor, , AltoColor, AnchoColor)
    Call SetBlock(900, 1, , "L")
    Call PrintData(color & "  " & cantidad)

    Call SetPos(120, 90)
    Call PrintBarCode(codigo, C90, 80, "N", "N")
    Call SetPos(120 + 90, 90 + 40)  'Relativo al Bar Code
    Call SetFont("B")
    Call PrintData(codigo)
                    
    Call SetFont
    Call SetPos(50, 20)
    Call SetBlock(975, , , "R")
    Call PrintData("HECHO EN PERU POR KALI")
                    
    Call SetFont
    Call SetPos(50, 50)
    Call SetBlock(975, , , "L")
    Call PrintData("UPO-FR004")

    '                'Imprimiendo Código de Barra y Logo
    '                Call SetPos(660, 30)
    '                    Call SetFont("F", , 120, 20)
    '                    Call SetBlock(308)
    '                    Call PrintData("CRESMAR")
    '                    'Call LoadImage(App.Path, "CRESMAR")
    '                    'Call PrintImage("CRESMAR", 2, 2)
            
    Call EndLabel
    Call ClosePrinter
    
End Sub

Private Sub SetNumOfCopies(Optional Number As Integer = 1)

    If Number > 1 Then
        NumOfCopies = "^PQ" & Str(Number)
    Else
        NumOfCopies = ""

    End If

End Sub

Private Sub SetPort(Optional strPort As String = "LPT1:")
    Port = strPort

End Sub

Private Sub SetPos(Y As Single, X As Single)  'Y NRO DE LINEA  X POSICION MAS A LA DERECHA

    Select Case mOrientation

        Case "N", "B"
            Call Add("^FO" & X & "," & Y)

        Case "R", "I"
            Call Add("^FO" & Y & "," & X)

    End Select
    
End Sub

Private Sub confCalibrate()
    Call Add("~JC")

End Sub

Private Sub confCalibrateShowGraphic()
    Call Add("~JG")

End Sub

Sub borrar_tablas(buf)

    On Error GoTo cmd89121_err

    cn.Execute (buf)
    Exit Sub
cmd89121_err:
    Exit Sub

End Sub

Private Sub confSetLabelLength()
    Call Add("~JL100")

End Sub

Private Sub Command10_Click()

    On Error GoTo cmd677_err

    Dim buf     As String

    Dim I       As Integer

    Dim bufg1   As String

    Dim xand    As String

    Dim xlistax As String

    If Len(tablas) = 0 Then Exit Sub
    If Combo4 = "%" Then
        MsgBox "Seleccione una Lista Precios", 48, "Aviso"
        Exit Sub

    End If

    buf = "Descripcio,Producto,"
    borrar_tablas "drop table prueba "
    buf = Trim(buf)
    buf = "select " + buf + " from " + tablas + " where familia like '" + extra_loquesea1(Combo1) & "'"

    If Len(Combo2) > 0 And Len(Combo3) > 0 And BUSCARE <> "%" Then
        buf = buf + " and " + Combo2 + Combo3 + BUSCARE

    End If

    buf = "select producto.selecciona,Producto.descripcio,producto.producto,Producto.Familia,Producto.descorto ,producto.costou,precios.pventa1 as Precio,precios.unidad1 as unidad,precios.factor1 as factor,Producto.Barras,producto.marca as Marca into prueba from producto left join precios on producto.producto=precios.producto  "

    If Combo1 <> "%" Then
        buf = buf & " where "
        buf = buf & "  producto.familia like '" + extra_loquesea1(Combo1) & "%'"

    End If

    If Combo2 <> "%" And Combo3 <> "%" And BUSCARE <> "%" Then
        If Combo1 <> "%" Then
            xand = " and "
        Else
            xand = " where "

        End If

        buf = buf + xand + Combo2 + Combo3 + "'" + BUSCARE & "'"

    End If

    buf = buf & " and precios.local='" & Combo4 & "'"
    buf = buf & " order by descripcio"
    'MsgBox buf
   
    cn.Execute (buf)

    If mytablexx.State = 1 Then mytablexx.Close
    mytablexx.Open "select  * from prueba ", cn, adOpenStatic, adLockOptimistic
    Set DBGrid2.DataSource = mytablexx
    DBGrid2.refresh

    DBGrid2.columns(0).Width = 600
    DBGrid2.columns(1).Width = 7500
    DBGrid2.columns(2).Width = 1000
    DBGrid2.columns(3).Width = 1000
    DBGrid2.columns(4).Width = 1000
    DBGrid2.columns(5).Width = 1000
    DBGrid2.columns(6).Width = 1000
    DBGrid2.columns(7).Width = 1000
    DBGrid2.columns(8).Width = 1000
    DBGrid2.columns(9).Width = 1000

    Frame3.Visible = True
    DBGrid2.SetFocus
    Exit Sub
cmd677_err:
    MsgBox "Intente de Nuevo " + error$, 24, "Aviso"
    Exit Sub

End Sub

Private Sub Command11_Click()

    On Error GoTo cmd89120_err

    dbzebra.Delete
    dbzebra.Requery
    dbGrid1.columns(0).Width = 600
    dbGrid1.columns(1).Width = 7500
    dbGrid1.columns(2).Width = 1000
    dbGrid1.columns(3).Width = 1000
    dbGrid1.columns(4).Width = 1000
    dbGrid1.columns(5).Width = 1000
    dbGrid1.columns(6).Width = 1000
    dbGrid1.columns(7).Width = 1000
    dbGrid1.columns(8).Width = 1000
    Exit Sub
cmd89120_err:
    Exit Sub

End Sub

Private Sub Command14_Click()
    Command10_Click

End Sub

Private Sub Command15_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

    If Len(Trim(diseno1)) = 0 Then Exit Sub
    guardaR_fichero

End Sub

Private Sub Command4_Click()

    On Error GoTo cmd34_err

    CmDialog1.InitDir = globaldir & "\zebra"
    CmDialog1.FileName = "*.*"
    CmDialog1.Filter = "(*.*)"
    'cmdialog1.FilterIndex = 1
    CmDialog1.Action = 1
    nombre_fichero = CmDialog1.FileName

    If nombre_fichero = "*.*" Then
        nombre_fichero = globaldir & "\zebra\DEMONIO"

    End If

    If Len(nombre_fichero) = 0 Then Exit Sub
    abrir_fichero
    Exit Sub
cmd34_err:
    Exit Sub

End Sub

Private Sub abrir_fichero()

    Dim buf As String

    Dim I, max

    On Error GoTo cmd23_err

    If Dir$(nombre_fichero) <> "" Then
        Open nombre_fichero For Input As #1
        buf = ""
        max = LOF(1)

        For I = 1 To max
            Seek #1, I
            buf = buf & input$(1, #1)
        Next I

        Close #1
        diseno1 = buf

    End If

    Exit Sub
cmd23_err:
    Close #1
    Exit Sub

End Sub

Private Sub guardaR_fichero()
    Call borrar_archivo
    Open nombre_fichero For Append As #5
    Print #5, diseno1
    Close #5
    Exit Sub

End Sub

Private Sub borrar_archivo()

    On Error GoTo cmd3_err

    Kill nombre_fichero
    Exit Sub
cmd3_err:
    Exit Sub

End Sub

Private Sub Command5_Click()
    Call imprimir_sticker

End Sub

Private Sub imprimir_datos()

    On Error GoTo cmd567891_err

    Dim linea$

    Dim buff$

    Dim campo      As String

    Dim j          As Integer

    Dim sw         As Integer

    Dim posicioni  As Long

    Dim posicionf  As Long

    Dim tlinea     As String

    Dim valor      As String

    Dim found      As Integer

    Dim nombrearch As String

    Dim posicionb  As Long

    Dim variable   As String

    Dim sw1        As Integer

    Dim Numero     As String

    Dim contando   As Integer

    posicionb = 1
    sw1 = 0
    Open nombre_fichero For Input As #7
    Do

        If EOF(7) Then Exit Do

        On Error GoTo cmd567891_err

        Line Input #7, buff

        On Error GoTo 0

        linea = Mid$(buff, 1, Len(buff))
        '-------------------------
        sw = 0
        posicioni = 0
        posicionf = 0
        valor = ""

        For j = 1 To Len(linea)

            If sw = 0 And Mid$(linea, j, 1) <> "[" And Mid$(linea, j, 1) <> "]" And Mid$(linea, j, 1) <> "{" And Mid$(linea, j, 1) <> "}" And Mid$(linea, j, 1) <> "/" And Mid$(linea, j, 1) <> "\" And Mid$(linea, j, 1) <> "<" And Mid$(linea, j, 1) <> ">" And Mid$(linea, j, 1) <> "^" And Mid$(linea, j, 1) <> "&" And Mid$(linea, j, 1) <> "(" And Mid$(linea, j, 1) <> ")" Then
                variable = Mid$(linea, j, 1)

                'found = formateaa(variable, 1, 0, 0)
            End If

            '------------------------------------------
            If Mid$(linea, j, 1) = "(" Then
                sw = 1
                posicioni = j + 1

            End If

            If sw = 1 And Mid$(linea, j, 1) = ")" Then
                posicionf = j - 1
                campo = Mid$(linea, posicioni, posicionf - posicioni + 1)
                valor = busca_campo1r(campo, Numero, contando)
                sw = 0
                posicioni = 0
                posicionf = 0

            End If

            If Mid$(linea, j, 1) = "[" Then
                sw = 1
                posicioni = j + 1

            End If

            If sw = 1 And Mid$(linea, j, 1) = "]" Then
                posicionf = j - 1
                campo = Mid$(linea, posicioni, posicionf - posicioni + 1)
                valor = busca_campo2(campo, Numero, contando)
                sw = 0
                posicioni = 0
                posicionf = 0

            End If

            '------------------------------------------
        Next j

        'found = formateaa("", 1, 2, 0)
        '-------------------------
    Loop
    Close #7
    
    Exit Sub
cmd567891_err:
    Close #7
    Exit Sub

End Sub

Private Sub imprimir_sticker()

    On Error GoTo cmd56789_err

    Dim linea$

    Dim buff$

    Dim campo      As String

    Dim j          As Integer

    Dim sw         As Integer

    Dim posicioni  As Long

    Dim posicionf  As Long

    Dim tlinea     As String

    Dim valor      As String

    Dim found      As Integer

    Dim nombrearch As String

    Dim posicionb  As Long

    Dim variable   As String

    Dim sw1        As Integer

    Dim Numero     As String

    Dim contando   As Integer

    If Val(numcopia) <= 0 Then Exit Sub
    Call SetPort("LPT1")
    Call SetOrientation(C0)
    Call OpenPrinter
    Call BeginLabel
    Call SetNumOfCopies(CInt(numcopia))
    posicionb = 1
    sw1 = 0
    Open nombre_fichero For Input As #7
    Do

        If EOF(7) Then Exit Do

        On Error GoTo cmd56789_err

        Line Input #7, buff

        On Error GoTo 0

        linea = Mid$(buff, 1, Len(buff))
        '-------------------------
        sw = 0
        posicioni = 0
        posicionf = 0
        valor = ""

        For j = 1 To Len(linea)

            If sw = 0 And Mid$(linea, j, 1) <> "[" And Mid$(linea, j, 1) <> "]" And Mid$(linea, j, 1) <> "{" And Mid$(linea, j, 1) <> "}" And Mid$(linea, j, 1) <> "/" And Mid$(linea, j, 1) <> "\" And Mid$(linea, j, 1) <> "<" And Mid$(linea, j, 1) <> ">" And Mid$(linea, j, 1) <> "^" And Mid$(linea, j, 1) <> "&" And Mid$(linea, j, 1) <> "(" And Mid$(linea, j, 1) <> ")" Then
                variable = Mid$(linea, j, 1)

                'found = formateaa(variable, 1, 0, 0)
            End If

            '------------------------------------------
            If Mid$(linea, j, 1) = "(" Then
                sw = 1
                posicioni = j + 1

            End If

            If sw = 1 And Mid$(linea, j, 1) = ")" Then
                posicionf = j - 1
                campo = Mid$(linea, posicioni, posicionf - posicioni + 1)
                valor = busca_campo1r(campo, Numero, contando)
                sw = 0
                posicioni = 0
                posicionf = 0

            End If

            '------------------------------------------
        Next j

        'found = formateaa("", 1, 2, 0)
        '-------------------------
    Loop
    Close #7
    Call EndLabel
    Call ClosePrinter
    
    Exit Sub
cmd56789_err:
    Close #7
    Call EndLabel
    Call ClosePrinter
    Exit Sub

End Sub

Function busca_campo2(campo As String, tablas As String, contando As Integer) As String

    Dim CAMPO1  As String

    Dim CAMPO2  As String

    Dim campo3  As String

    Dim campo4  As String

    Dim campo5  As String

    Dim campo6  As String

    Dim campo7  As String

    Dim campo8  As String

    Dim campo9  As String

    Dim campo10 As String

    Dim found   As Integer

    Dim buf     As String

    Dim bufy    As String

    Dim bufz    As String

    buf = campo
    CAMPO1 = ""
    CAMPO2 = ""
    campo3 = ""
    campo4 = ""
    campo5 = ""
    campo6 = ""
    campo7 = ""
    campo8 = ""
    campo9 = ""
    campo10 = ""

    If InStr(campo, ",") > 0 Then   'si es comna
        found = extraer_camposs(buf, CAMPO1, CAMPO2, campo3, campo4, campo5, campo6, campo7, campo8, campo9, campo10, ",")
        'imprimir formatos
        '------------------------------------------------
        'MsgBox campo1 & " " & campo2 & " " & campo3 & " " & campo4 & " " & campo5 & " " & campo6 & " " & campo7 & " " & campo8 & " " & campo9
        'MsgBox posicion
        CAMPO1 = Trim(CAMPO1)
        CAMPO2 = Trim(CAMPO2)
        campo3 = Trim(campo3)
        campo4 = Trim(campo4)
        campo5 = Trim(campo5)
        campo6 = Trim(campo6)
        campo7 = Trim(campo7)
        campo8 = Trim(campo8)
        campo9 = Trim(campo9)
        campo10 = Trim(campo10)

        If Val(campo8) = 0 Then
            campo8 = "1"

        End If

        If Val(campo9) = 0 Then
            campo9 = "1"

        End If

        If CAMPO2 = "TEXTO" Then
            Call SetPos(CSng(campo3), CSng(campo4) + posicion)
            Call SetFont(campo5, , CSng(campo6), CSng(campo7))
            Call SetBlock(700, , , "L")
            'MsgBox "" & dbzebra.Fields(campo1)
            bufy = Mid$("" & dbzebra.Fields(CAMPO1), Val(campo8), Val(campo9))
            bufy = Trim(bufy)

            'MsgBox CAMPO1 & " " & dbzebra.Fields(CAMPO1).Type
            If "" & dbzebra.Fields(CAMPO1).Type = "5" Then  'FLOAT
                bufy = Format(Val(bufy), "0.00")

            End If

            'MsgBox "1." & bufy & " " & Len(bufy) & "  " & campo9
            If campo10 = "C" Then
                bufy = centrar_campos(bufy, Val(campo9))

            End If

            'MsgBox "2." & bufz & " " & Len(bufz) & "  " & campo9
            Call PrintData(bufy)

        End If

        If CAMPO2 = "BARRA" Then
            Call SetPos(CSng(campo3), CSng(campo4) + posicion)
            Call PrintBarCode("" & dbzebra.Fields(CAMPO1), , 50, "N", "N")

        End If

        '------------------------------------------------
    End If

End Function

Function centrar_campos(buf As String, longitud As Integer) As String

    Dim buff As String

    Dim I    As Integer

    Dim X    As Integer

    buff = "."
    X = 0

    If Len(buf) < longitud Then
        X = (longitud - Len(buf)) / 2
        buff = buff & String(X, Chr$(32)) + buf
    Else
        buff = buf

    End If

    centrar_campos = buff

End Function

Function busca_campo1r(campo As String, tablas As String, contando As Integer) As String

    Dim CAMPO1  As String

    Dim CAMPO2  As String

    Dim campo3  As String

    Dim campo4  As String

    Dim campo5  As String

    Dim campo6  As String

    Dim campo7  As String

    Dim campo8  As String

    Dim campo9  As String

    Dim campo10 As String

    Dim found   As Integer

    Dim buf     As String

    buf = campo
    CAMPO1 = ""
    CAMPO2 = ""
    campo3 = ""
    campo4 = ""
    campo5 = ""
    campo6 = ""
    campo7 = ""
    campo8 = ""
    campo9 = ""
    campo10 = ""

    If InStr(campo, ",") > 0 Then   'si es comna
        found = extraer_camposs(buf, CAMPO1, CAMPO2, campo3, campo4, campo5, campo6, campo7, campo8, campo9, campo10, ",")
        'imprimir formatos
        '------------------------------------------------
        'MsgBox campo1 & " " & campo2 & " " & campo3 & " " & campo4 & " " & campo5 & " " & campo6 & " " & campo7 & " " & campo8 & " " & campo9
        CAMPO1 = Trim(CAMPO1)
        CAMPO2 = Trim(CAMPO2)
        campo3 = Trim(campo3)
        campo4 = Trim(campo4)
        campo5 = Trim(campo5)
        campo6 = Trim(campo6)
        campo7 = Trim(campo7)
        campo8 = Trim(campo8)
        campo9 = Trim(campo9)
        campo10 = Trim(campo10)

        If CAMPO2 = "TEXTO" Then
            Call SetPos(CSng(campo3), CSng(campo4))
            Call SetFont(campo5, , CSng(campo6), CSng(campo7))
            Call SetBlock(700, , , "L")
   
            Call PrintData(Mid$(CAMPO1, CSng(campo8), CSng(campo9)))

        End If

        If CAMPO2 = "BARRA" Then
            Call SetPos(CSng(campo3), CSng(campo4))
            Call PrintBarCode(CAMPO1, , 50, "N", "N")

        End If

        '------------------------------------------------
    End If

End Function

Function extraer_camposs(campo As String, _
                         CAMPO1 As String, _
                         CAMPO2 As String, _
                         campo3 As String, _
                         campo4 As String, _
                         campo5 As String, _
                         campo6 As String, _
                         campo7 As String, _
                         campo8 As String, _
                         campo9 As String, _
                         campo10 As String, _
                         Flags As String)

    Dim I    As Integer

    Dim j    As Integer

    Dim temp As String

    I = 0
    temp = Trim$(campo)

    If Len(temp) = 0 Then Exit Function
    Do
        j = InStr(temp, Flags)

        If j > 0 Then
            I = I + 1

            Select Case I

                Case 1: CAMPO1 = Mid$(temp, 1, j - 1)

                Case 2: CAMPO2 = Mid$(temp, 1, j - 1)

                Case 3: campo3 = Mid$(temp, 1, j - 1)

                Case 4: campo4 = Mid$(temp, 1, j - 1)

                Case 5: campo5 = Mid$(temp, 1, j - 1)

                Case 6: campo6 = Mid$(temp, 1, j - 1)

                Case 7: campo7 = Mid$(temp, 1, j - 1)

                Case 8: campo8 = Mid$(temp, 1, j - 1)

                Case 9: campo9 = Mid$(temp, 1, j - 1)

                Case 10: campo10 = Mid$(temp, 1, j - 1)
            
            End Select

            temp = Trim$(Mid$(temp, j + 1))
        Else
            Exit Function

        End If

    Loop
    Exit Function

End Function

Private Sub Command6_Click()

End Sub

Private Sub Command7_Click()

    Dim archivo_name As String

    'archivo_name = nombre_fichero
    archivo_name = InputBox("Guardar Archivo", "Guardar como..", "")

    If Len(archivo_name) <= 8 And Len(archivo_name) >= 1 Then
        nombre_fichero = globaldir & "\zebra\" & archivo_name
        guardaR_fichero

    End If

End Sub

Private Sub Command8_Click()
    Call trescolumnas

End Sub

Private Sub Command9_Click()

    refresca_db

End Sub

Private Sub dki1212_Click()
    Frame1.Visible = True
    diseno1.SetFocus
    Exit Sub

End Sub

Private Sub dbgrid1_BeforeColUpdate(ByVal ColIndex As Integer, _
                                    OldValue As Variant, _
                                    Cancel As Integer)

    Select Case ColIndex

        Case 0

            'MsgBox dbgrid1.columns(0)
            'If Not IsNumeric(dbgrid1.columns(0)) Then
            '    MsgBox "Dato no Numerico", 24, "Aviso"
            '    Cancel = True
            '    Exit Sub
            'End If
    End Select

End Sub

Private Sub dbgrid2_AfterColUpdate(ByVal ColIndex As Integer)
    ' dbzebra.Update

End Sub

Private Sub DBGrid2_DblClick()

    Dim found As Integer

    found = verifica_existe("" & Trim(mytablexx.Fields("producto")))

    If found = 1 Then
        MsgBox "Ya existe Seleccionado", 24, "Aviso"
        Exit Sub

    End If

    If "" & mytablexx.Fields("Selecciona") = "S" Then
        mytablexx.Fields("Selecciona") = ""
        mytablexx.Update
        Exit Sub

    End If

    If Trim("" & mytablexx.Fields("Selecciona")) <> "S" Then
        mytablexx.Fields("Selecciona") = "S"
        mytablexx.Update
        Exit Sub

    End If

    mytablexx.Update
    Exit Sub

    'Call agregar_seleccion
End Sub

Private Sub dbgrid2_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim found As Integer

    On Error GoTo cmd23_err

    If KeyCode = 13 Then
        found = verifica_existe("" & Trim(mytablexx.Fields("producto")))

        If found = 1 Then
            MsgBox "Ya existe Seleccionado", 24, "Aviso"
            Exit Sub

        End If
   
        Call agregar_seleccion

    End If

    If KeyCode = 27 Then
        Frame3.Visible = False
        Exit Sub

    End If

    Exit Sub
cmd23_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub agregar_seleccion()

    Dim I As Integer

    On Error GoTo cmd89111_err

    dbzebra.AddNew
    
    dbzebra.Fields("PRODUCTO") = "" & mytablexx.Fields("producto")
    dbzebra.Fields("descripcio") = "" & mytablexx.Fields("descripcio")
    dbzebra.Fields("precio") = Val(Format(Val("" & mytablexx.Fields("precio")), "0.00"))
    dbzebra.Fields("barras") = "" & mytablexx.Fields("barras")
    dbzebra.Fields("descorto") = "" & mytablexx.Fields("descorto")
    dbzebra.Fields("familia") = "" & mytablexx.Fields("familia")
    dbzebra.Fields("marca") = "" & mytablexx.Fields("marca")
    dbzebra.Fields("unidad1") = "" & mytablexx.Fields("unidad1")
    dbzebra.Fields("acanti") = 1
    
    dbzebra.Update
    Exit Sub
cmd89111_err:
    MsgBox "Aviso en agregar Seleccion " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Load()

    Dim mytablex As New ADODB.Recordset

    Dim I        As Integer

    On Error GoTo cmd34_err

    nombre_fichero = globaldir & "\zebra\demonio"
    Combo4.Clear
    Combo4.AddItem "%"

    For I = 1 To 9
        Combo4.AddItem Format(I, "00")
    Next I

    Combo4.ListIndex = 0
    Combo3.Clear
    Combo3.AddItem "%"
    Combo3.AddItem " LIKE "
    Combo3.AddItem " > "
    Combo3.AddItem " < "
    Combo3.AddItem " >= "
    Combo3.AddItem " <= "
    Combo3.AddItem " = "
    Combo3.ListIndex = 0
   
    Combo1.Clear
    mytablex.Open "select * from familia order by descripcio", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        Combo1.AddItem Trim("" & mytablex.Fields("descripcio")) & "|" & Trim("" & mytablex.Fields("familia"))
        mytablex.MoveNext
    Loop
    Combo1.AddItem "%"
    mytablex.Close

    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "producto.Descripcio"
    Combo2.AddItem "producto.Producto"
    Combo2.AddItem "producto.Barras"
    Combo2.AddItem "producto.Familia"
    Combo2.AddItem "precios.Unidad1"
    Combo2.ListIndex = 0
    borrar_zebra
    Command9_Click
    Exit Sub
cmd34_err:
    Exit Sub

End Sub

Private Sub cambia_grid()

End Sub

Function verifica_existe(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM ZEBRA WHERE producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        verifica_existe = 1

    End If

    mytablex.Close

End Function

Private Sub Image1_Click()
    Frame3.Visible = False
    dbGrid1.SetFocus
    Exit Sub

End Sub

Private Sub Image2_Click()

    If Frame4.Visible = True Then Exit Sub
    Command7.Enabled = True
    Command3.Enabled = True
    diseno1.Enabled = True
    Frame1.Visible = True
    abrir_fichero
    diseno1.SetFocus

End Sub

Private Sub Image7_Click()

    On Error GoTo cmd35_err

    CmDialog1.InitDir = globaldir & "\zebra"
    CmDialog1.FileName = "*.*"
    CmDialog1.Filter = "(*.*)"
    'cmdialog1.FilterIndex = 1
    CmDialog1.Action = 1
    nombre_fichero = CmDialog1.FileName

    If nombre_fichero = "*.*" Then
        nombre_fichero = globaldir & "\zebra\demonio"

    End If

    abrir_fichero
    Exit Sub
cmd35_err:

End Sub

Private Sub Image8_Click()

    If Frame4.Visible = True Then Exit Sub
    Frame1.Visible = True
    abrir_fichero
    Command7.Enabled = False
    Command3.Enabled = False
    diseno1.Enabled = False
    numcopia.SetFocus

End Sub

Private Sub Label6_Click()

    'Frame2.Visible = True
    'List2.Clear
    'List1.SetFocus
End Sub

Private Sub Image3_Click()

    If Frame1.Visible = True Then Exit Sub
    Frame4.Visible = True
    dbGrid1.SetFocus

End Sub

Private Sub Image5_Click()
    Call ldso232_Click

End Sub

Private Sub Label1_Click()

    'Frame2.Visible = True
End Sub

Private Sub Label8_Click()

    On Error GoTo cmd9012_err

    If mytablexx.RecordCount = 0 Then Exit Sub
    mytablexx.MoveFirst
    Do

        If mytablexx.EOF Then Exit Do
        If "" & mytablexx.Fields("selecciona") = "S" Then
            copiandox

        End If

        mytablexx.MoveNext
    Loop
    Frame3.Visible = False
    dbGrid1.SetFocus
    Exit Sub
cmd9012_err:
    MsgBox "Aviso en copiar..." & error$, 48, "Aviso"
    Exit Sub

End Sub

Sub copiandox()

    On Error GoTo cmd4545_err

    dbzebra.AddNew
    dbzebra.Fields("PRODUCTO") = "" & mytablexx.Fields("producto")
    dbzebra.Fields("descripcio") = "" & mytablexx.Fields("descripcio")
    dbzebra.Fields("precio") = Val(Format(Val("" & mytablexx.Fields("precio")), "0.00"))
    dbzebra.Fields("barras") = "" & mytablexx.Fields("barras")
    dbzebra.Fields("descorto") = "" & mytablexx.Fields("descorto")
    dbzebra.Fields("familia") = "" & mytablexx.Fields("familia")
    dbzebra.Fields("marca") = "" & mytablexx.Fields("marca")
    dbzebra.Fields("unidad") = "" & mytablexx.Fields("unidad")
    dbzebra.Fields("acanti") = 1
    dbzebra.Update
    Exit Sub
cmd4545_err:
    MsgBox "Aviso en copiand " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub ldso232_Click()

    If Frame3.Visible = True Then
        Frame3.Visible = False
        Exit Sub

    End If

    If Frame1.Visible = True Then
        Frame1.Visible = False
        Exit Sub

    End If

    If Frame4.Visible = True Then
        Frame4.Visible = False
        Frame3.Visible = False
        Exit Sub

    End If

    xzebra.Hide
    Unload xzebra

End Sub

Private Sub List1_DblClick()

End Sub

Sub refresca_db()

    Dim buf1 As String

    Dim buf  As String

    If dbzebra.State = 1 Then
        dbGrid1.refresh
        Exit Sub

    End If
   
    buf1 = "Acanti,Descripcio,Producto,Precio,Unidad,Familia,Subfamilia,Marca,Unidad,Descorto,barras"
    buf = "select " & buf1 & " from zebra "
    'MsgBox buf
    dbzebra.Open buf, cn, adOpenDynamic, adLockOptimistic
    Set dbGrid1.DataSource = dbzebra
    dbGrid1.refresh
    dbGrid1.columns(0).Width = 600
    dbGrid1.columns(1).Width = 7500
    dbGrid1.columns(2).Width = 1000
    dbGrid1.columns(3).Width = 1000
    dbGrid1.columns(4).Width = 1000
    dbGrid1.columns(5).Width = 1000
    dbGrid1.columns(6).Width = 1000
    dbGrid1.columns(7).Width = 1000
    dbGrid1.columns(8).Width = 1000
    'dbgrid1.columns(9).Width = 1000

End Sub

Sub borrar_zebra()
    borrar_tablas "drop table zebra"
    borrar_tablas "select * into zebra from producto"
    borrar_tablas "delete from zebra"

End Sub

Private Sub List2_Click()

End Sub
