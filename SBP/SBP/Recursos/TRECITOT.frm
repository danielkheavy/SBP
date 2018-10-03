VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form trecitot 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comprobantes Ingresos Egresos"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   15225
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox subconcepto 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   48
      Top             =   7440
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.ComboBox concepto 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Top             =   7080
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      Caption         =   "Consulta"
      ForeColor       =   &H8000000B&
      Height          =   3615
      Left            =   1560
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   7335
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
         Height          =   735
         Left            =   5760
         MaskColor       =   &H00FFFFFF&
         Picture         =   "TRECITOT.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   1335
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
         Left            =   5760
         MaskColor       =   &H00FFFFFF&
         Picture         =   "TRECITOT.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00808080&
      Height          =   1155
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   15165
      TabIndex        =   0
      Top             =   0
      Width           =   15225
      Begin VB.ComboBox tacu 
         BackColor       =   &H00FFFFFF&
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
         Left            =   11160
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   0
         Width           =   2415
      End
      Begin VB.ComboBox local1 
         BackColor       =   &H00FFFFFF&
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
         Left            =   11160
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   720
         Width           =   2415
      End
      Begin VB.ComboBox tipoclie 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   0
         Width           =   615
      End
      Begin VB.ComboBox turno 
         BackColor       =   &H00FFFFFF&
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
         Left            =   8640
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox tipo 
         BackColor       =   &H00FFFFFF&
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
         Left            =   11160
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   14640
         MaskColor       =   &H00FFFFFF&
         Picture         =   "TRECITOT.frx":0F5C
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   600
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cajero 
         BackColor       =   &H00FFFFFF&
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
         Left            =   8640
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox codigo 
         Height          =   375
         Left            =   6840
         MaxLength       =   11
         TabIndex        =   18
         Text            =   "%"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox nombre 
         Height          =   375
         Left            =   6840
         MaxLength       =   11
         TabIndex        =   17
         Text            =   "%"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox fechai 
         BackColor       =   &H0080FFFF&
         Height          =   375
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   14
         Top             =   0
         Width           =   1815
      End
      Begin VB.TextBox fechaf 
         BackColor       =   &H0080FFFF&
         Height          =   375
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   13
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox caja 
         BackColor       =   &H00FFFFFF&
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
         Left            =   8640
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   0
         Width           =   1815
      End
      Begin VB.CommandButton cmdSort 
         BackColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TRECITOT.frx":170A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Consulta"
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
         Left            =   1440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TRECITOT.frx":291C
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Left            =   0
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TRECITOT.frx":3B2E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprimir"
         Top             =   0
         Width           =   735
      End
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn1 
         Height          =   810
         Left            =   13680
         TabIndex        =   49
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1429
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
         MICON           =   "TRECITOT.frx":4D40
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "W/V"
         Height          =   255
         Left            =   10560
         TabIndex        =   35
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ClieTipo"
         Height          =   375
         Left            =   6000
         TabIndex        =   31
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Turno"
         Height          =   495
         Left            =   8040
         TabIndex        =   29
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
         Height          =   375
         Left            =   10560
         TabIndex        =   27
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         Height          =   255
         Left            =   10560
         TabIndex        =   26
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Usuario"
         Height          =   375
         Left            =   8040
         TabIndex        =   22
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         Height          =   375
         Left            =   6000
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         Height          =   375
         Left            =   6000
         TabIndex        =   19
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   375
         Left            =   3120
         TabIndex        =   16
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   375
         Left            =   3120
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja"
         Height          =   375
         Left            =   8040
         TabIndex        =   12
         Top             =   0
         Width           =   615
      End
   End
   Begin MSDataGridLib.DataGrid dbgrid2 
      Height          =   5775
      Left            =   0
      TabIndex        =   32
      Top             =   1080
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   10186
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
      ColumnCount     =   19
      BeginProperty Column00 
         DataField       =   "Estado"
         Caption         =   "E"
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
         DataField       =   "Local"
         Caption         =   "Local"
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
         DataField       =   "Tipo"
         Caption         =   "Tipo"
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
         DataField       =   "Serie"
         Caption         =   "Serie"
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
      BeginProperty Column04 
         DataField       =   "Numero"
         Caption         =   "Numero"
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
      BeginProperty Column05 
         DataField       =   "fecha"
         Caption         =   "Fecha"
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
         DataField       =   "Hora"
         Caption         =   "Hora"
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
      BeginProperty Column07 
         DataField       =   "Tipoclie"
         Caption         =   "T"
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
      BeginProperty Column08 
         DataField       =   "Codigo"
         Caption         =   "Codigo"
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
      BeginProperty Column09 
         DataField       =   "Nombre"
         Caption         =   "Nombre"
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
      BeginProperty Column10 
         DataField       =   "Moneda"
         Caption         =   "M"
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
      BeginProperty Column11 
         DataField       =   "Total"
         Caption         =   "Total"
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
      BeginProperty Column12 
         DataField       =   "Usuario"
         Caption         =   "Cajero"
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
      BeginProperty Column13 
         DataField       =   "Caja"
         Caption         =   "Caja"
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
      BeginProperty Column14 
         DataField       =   "Turno"
         Caption         =   "T"
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
      BeginProperty Column15 
         DataField       =   "vendedor"
         Caption         =   "Vendedor"
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
      BeginProperty Column16 
         DataField       =   "Acu"
         Caption         =   "Acu"
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
      BeginProperty Column17 
         DataField       =   "Afecta"
         Caption         =   "A"
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
      BeginProperty Column18 
         DataField       =   "Observa"
         Caption         =   "Observa"
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
         MarqueeStyle    =   2
         BeginProperty Column00 
            ColumnWidth     =   225.071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   180.283
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   2009.764
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   225.071
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   390.047
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   195.024
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   269.858
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   2234.835
         EndProperty
      EndProperty
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SubConcepto"
      Height          =   375
      Left            =   120
      TabIndex        =   47
      Top             =   7440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Concepto"
      Height          =   375
      Left            =   120
      TabIndex        =   45
      Top             =   7080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo"
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
      Left            =   8400
      TabIndex        =   44
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label esoles 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11640
      TabIndex        =   43
      Top             =   7800
      Width           =   1695
   End
   Begin VB.Label edolares 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10080
      TabIndex        =   42
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Egresos"
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
      Left            =   8400
      TabIndex        =   41
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label isoles 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11640
      TabIndex        =   40
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label idolares 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10080
      TabIndex        =   39
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ingresos"
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
      Left            =   8400
      TabIndex        =   38
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Soles"
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
      Left            =   11640
      TabIndex        =   37
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dolares"
      Height          =   375
      Left            =   10080
      TabIndex        =   36
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label estaya 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   13680
      TabIndex        =   24
      Top             =   7920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label dolares 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10080
      TabIndex        =   10
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label soles 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11640
      TabIndex        =   9
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Label afecta 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   13680
      TabIndex        =   8
      Top             =   8160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label acu 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   14400
      TabIndex        =   7
      Top             =   7920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Menu dnu823 
      Caption         =   "&Nuevo"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu dbo912 
      Caption         =   "&Borra"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu dki9923 
      Caption         =   "&Consulta"
      Visible         =   0   'False
   End
   Begin VB.Menu dfkl8823 
      Caption         =   "&Copia"
   End
   Begin VB.Menu dkj8933 
      Caption         =   "&Reporte"
   End
   Begin VB.Menu lfo3434 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "trecitot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rconsulta As New ADODB.Recordset

Private Sub ChameleonBtn1_Click()
    sql_recibos

End Sub

Private Sub cmdCancelar_Click()
    lfo3434_Click

End Sub

Private Sub cmdDelete_Click()

    'dbo912_Click
End Sub

Private Sub cmdExit_Click()

    If Frame2.Visible = True Then
        Frame2.Visible = False
        Exit Sub

    End If

    trecitot.Hide
    Unload trecitot

End Sub

Private Sub cmdGrabar_Click()
    sql_recibos
    lfo3434_Click

End Sub

Function descarga_cuentac(xlocal1 As String, _
                          xtipo1 As String, _
                          xserie1 As String, _
                          xnumero1 As String, _
                          signo As String)

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim buf      As String

    Dim buf1     As String

    On Error GoTo cmd43_err

    '----primero debe habecerse grabdo el temporal y luego seleccionar para descargar
    If "" & DBGrid2.columns("t") = "C" Then
        buf = "cuentac"
        buf1 = "cuentacd"

    End If

    If "" & DBGrid2.columns("t") = "P" Then
        buf = "cuentap"
        buf1 = "cuentapd"

    End If

    If "" & DBGrid2.columns("t") = "V" Then
        buf = "cuentac"
        buf1 = "cuentacd"

    End If

    mytabley.Open "select * from " & buf1 & " where local='" & xlocal1 & "' and tipo='" & xtipo1 & "' and serie='" & xserie1 & "' and numero='" & xnumero1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        Do

            If mytabley.EOF Then Exit Do
            mytablex.Open "select * from " & buf & " where local='" & "" & mytabley.Fields("local") & "' and tipo='" & "" & mytabley.Fields("tipo1") & "' and serie='" & "" & mytabley.Fields("serie1") & "' and numero='" & "" & mytabley.Fields("numero1") & "' and cuota='" & "" & mytabley.Fields("cuota") & "'", cn, adOpenStatic, adLockOptimistic

            If mytablex.RecordCount > 0 Then
                'mytablex.Edit
                sdx = Val("" & mytablex.Fields("abono")) + Val(signo) * Val("" & mytabley.Fields("paga"))
                mytablex.Fields("abono") = Val(Format(sdx, "0.00"))
                sdx = Val("" & mytablex.Fields("total")) - Val("" & mytablex.Fields("abono")) + Val("" & mytablex.Fields("interes"))
                mytablex.Fields("saldo") = Val(Format(sdx, "0.00"))
                sdx = Val("" & mytablex.Fields("c1")) - Val(signo) * Val("" & mytabley.Fields("l1"))
                mytablex.Fields("c1") = Format(sdx, "0.00")
                sdx = Val("" & mytablex.Fields("c2")) - Val(signo) * Val("" & mytabley.Fields("l2"))
                mytablex.Fields("c2") = Format(sdx, "0.00")
                sdx = Val("" & mytablex.Fields("c3")) - Val(signo) * Val("" & mytabley.Fields("l3"))
                mytablex.Fields("c3") = Format(sdx, "0.00")
                sdx = Val("" & mytablex.Fields("c4")) - Val(signo) * Val("" & mytabley.Fields("l4"))
                mytablex.Fields("c4") = Format(sdx, "0.00")
                mytablex.Update

            End If

            mytablex.Close
            mytabley.MoveNext
        Loop

    End If

    '----ahora lo borramos tmpcta-----
    cn.Execute ("delete from " & buf1 & " where local='" & xlocal1 & "' and tipo='" & xtipo1 & "' and serie='" & xserie1 & "' and numero='" & xnumero1 & "'")
    mytabley.Close
    Exit Function
cmd43_err:
    MsgBox "Aviso en descarga cuentac ", 48, "Aviso"
    Exit Function

End Function

Function descarga_letra(xlocal1 As String, _
                        xtipo1 As String, _
                        serie1 As String, _
                        xnumero1 As String, _
                        signo As String)

    Dim sdx      As Double

    Dim buf      As String

    Dim buf1     As String

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    If Len(xnumero1) = 0 Then Exit Function

    'If Len(xtotal1) = 0 Then Exit Function
    If "" & DBGrid2.columns("tipoclie") = "C" Then
        buf = "letrav"
        buf1 = "letracd"

    End If

    If "" & DBGrid2.columns("tipoclie") = "P" Then
        buf = "letrac"
        buf1 = "letrapd"

    End If

    If "" & DBGrid2.columns("tipoclie") = "V" Then
        buf = "letrav"
        buf1 = "letracd"

    End If

    mytabley.Open "select * FROM " & buf1 & " where local='" & xlocal1 & "' and tipo='" & xtipo1 & "' and serie='" & xserie1 & "' and numero='" & xnumero1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        Do

            If mytabley.EOF Then Exit Do
            mytablex.Open "select * " & buf & " where local='" & "" & mytabley.Fields("local") & "' and tipo='" & "" & mytabley.Fields("tipo1") & "' and serie='" & "" & mytabley.Fields("serie1") & "' and numero='" & "" & mytabley.Fields("numero1") & "' and cuota='" & "" & mytabley.Fields("cuota") & "'", cn, adOpenStatic, adLockOptimistic

            If mytablex.RecordCount > 0 Then
                'mytablex.Edit
                sdx = Val("" & mytablex.Fields("abono")) + Val(signo) * Val("" & mytabley.Fields("paga"))
                mytablex.Fields("abono") = Val(Format(sdx, "0.00"))
                sdx = Val("" & mytablex.Fields("importe")) - Val("" & mytablex.Fields("amortiza")) + Val("" & mytablex.Fields("interes1")) + Val("" & mytablex.Fields("interes2")) + Val("" & mytablex.Fields("protesto")) + Val("" & mytablex.Fields("otros")) - Val("" & mytablex.Fields("abono"))
                mytablex.Fields("saldo") = Val(Format(sdx, "0.00"))
   
                sdx = Val("" & mytablex.Fields("c1")) - Val(signo) * Val("" & mytabley.Fields("l1"))
                mytablex.Fields("c1") = Format(sdx, "0.00")
                sdx = Val("" & mytablex.Fields("c2")) - Val(signo) * Val("" & mytabley.Fields("l2"))
                mytablex.Fields("c2") = Format(sdx, "0.00")
                sdx = Val("" & mytablex.Fields("c3")) - Val(signo) * Val("" & mytabley.Fields("l3"))
                mytablex.Fields("c3") = Format(sdx, "0.00")
                sdx = Val("" & mytablex.Fields("c4")) - Val(signo) * Val("" & mytabley.Fields("l4"))
                mytablex.Fields("c4") = Format(sdx, "0.00")
                mytablex.Update

            End If

            mytablex.Close
            mytabley.MoveNext
        Loop

    End If

    mytabley.Close
 
End Function

Private Sub cmdPrint_Click()
    dkj8933_Click

    'repingre.acu = acu
    'repingre.Show 1
End Sub

Private Sub Command1_Click()
    sql_recibos

End Sub

Private Sub concepto_Click()

    If concepto = "%" Then Exit Sub
    carga_subconcepto extra_loquesea(concepto)

End Sub

Private Sub dbo912_Click()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd34_err

    If DBGrid2.columns("e") <> "2" Then
        MsgBox "Debe estar en estado 0", 1, "Aviso"
        Exit Sub

    End If

    If MsgBox("Desea Borra el registro " & DBGrid2.columns("numero"), 1, "Aviso") <> "1" Then Exit Sub
    If "" & DBGrid2.columns("acu") = "W" Or "" & DBGrid2.columns("acu") = "V" Then  'ingreso/egreso
        found = descarga_cuentac("" & DBGrid2.columns("local"), "" & DBGrid2.columns("tipo"), "" & DBGrid2.columns("serie"), "" & DBGrid2.columns("numero"), "-1")

    End If

    found = borra_fpagov("" & DBGrid2.columns("local"), "" & DBGrid2.columns("tipo"), "" & DBGrid2.columns("serie"), "" & DBGrid2.columns("numero"))
    cn.Execute ("delete from recibo where local='" & DBGrid2.columns("local") & "' and tipo='" & "" & DBGrid2.columns("tipo") & "' and serie='" & "" & DBGrid2.columns("serie") & "' and numero='" & "" & DBGrid2.columns("numero") & "'")
    sql_recibos
    MsgBox "Proceso Borrado ", 48, "Aviso"
    Exit Sub
cmd34_err:
    MsgBox "Aviso en borra " + error$, 48, "Aviso"
    Exit Sub

End Sub

Function borra_fpagov(xlocal As String, _
                      xtipo As String, _
                      xserie As String, _
                      xnumero As String)
    cn.Execute ("delete from fpagov where local='" & xlocal & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'")

End Function

Private Sub dfkl8823_Click()

    On Error GoTo cdm99_err

    proceso_impresion1 "" & DBGrid2.columns("t"), "" & DBGrid2.columns("local"), "" & DBGrid2.columns("tipo"), "" & DBGrid2.columns("serie"), "" & DBGrid2.columns("numero"), "" & DBGrid2.columns("acu")
    Exit Sub
cdm99_err:
    MsgBox "Seleccione un registro", 48, "Aviso"
    Exit Sub
   
End Sub

Private Sub dki9923_Click()
    Frame2.Visible = True
    fechai.SetFocus

End Sub

'' 30/11/2017 Mejora reporte ingresos/egresos

'Private Sub dkj8933_Click()
'
' Dim v, h As Long
' Dim found As Integer
' Dim i As Integer
' Dim R As Long
' Dim sdx As Double
' Dim sdx1 As Double
' Dim sdx2 As Double
' Dim sdx3 As Double
' Dim sdx4 As Double
' Dim xingreso As Double
' Dim xegreso As Double
'
'    Dim Heading(12) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion
'    On Error GoTo cmd1561212_err
'    If MsgBox("Desea Generar reporte ", 1, "Aviso") <> 1 Then Exit Sub
'    If rconsulta.RecordCount = 0 Then Exit Sub
'    rconsulta.MoveFirst
'
'    Heading(1) = "Lo"
'    Heading(2) = "Tipo"
'    Heading(3) = "Serie"
'    Heading(4) = "Numero"
'    Heading(5) = "Fecha"
'    Heading(6) = "Codigo"
'    Heading(7) = "Nombre"
'    Heading(8) = "M"
'    Heading(9) = "Ingreso"
'    Heading(10) = "Egreso"
'
'
'    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
'    Call Formato_Excelre(12, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
'
'v = 5
'h = 1
'sdx = 0
'sdx1 = 0
'sdx2 = 0
'sdx3 = 0
'sdx4 = 0
'
'    objExcel.ActiveSheet.Cells(v, h + 1) = "Reporte de ingresos Egresos"
'    v = v + 1
'    objExcel.ActiveSheet.Cells(v, h + 1) = "FechaI:" & fechai & " Fechaf:" & fechaf
'    v = v + 1
'
'     Do
'            If rconsulta.EOF Then Exit Do
'            objExcel.ActiveSheet.Cells(v, h) = "'" & rconsulta.Fields("local")
'            objExcel.ActiveSheet.Cells(v, h + 1) = "'" & rconsulta.Fields("tipo")
'            objExcel.ActiveSheet.Cells(v, h + 2) = "'" & rconsulta.Fields("serie")
'            objExcel.ActiveSheet.Cells(v, h + 3) = "'" & rconsulta.Fields("numero")
'            objExcel.ActiveSheet.Cells(v, h + 4) = "'" & rconsulta.Fields("fecha")
'            objExcel.ActiveSheet.Cells(v, h + 5) = "'" & rconsulta.Fields("codigo")
'            objExcel.ActiveSheet.Cells(v, h + 6) = "'" & rconsulta.Fields("nombre")
'            objExcel.ActiveSheet.Cells(v, h + 7) = "'" & rconsulta.Fields("moneda")
'            xingreso = 0
'            xegreso = 0
'            If Trim("" & rconsulta.Fields("acu")) = "W" Then
'            xingreso = Val("" & rconsulta.Fields("total"))
'            End If
'            If Trim("" & rconsulta.Fields("acu")) = "V" Then
'            xegreso = Val("" & rconsulta.Fields("total"))
'            End If
'            objExcel.ActiveSheet.Cells(v, h + 8) = xingreso
'            objExcel.ActiveSheet.Cells(v, h + 9) = xegreso
'            objExcel.ActiveSheet.Cells(v, h + 10) = "'" & rconsulta.Fields("observa")
'
'            v = v + 1
'            If Trim("" & rconsulta.Fields("moneda")) = "S" Then
'            If Trim("" & rconsulta.Fields("acu")) = "W" Then
'            sdx1 = sdx1 + Val("" & rconsulta.Fields("total"))
'            End If
'            If Trim("" & rconsulta.Fields("acu")) = "V" Then
'            sdx2 = sdx2 + Val("" & rconsulta.Fields("total"))
'            End If
'            End If
'
'            If Trim("" & rconsulta.Fields("moneda")) = "D" Then
'            If Trim("" & rconsulta.Fields("acu")) = "W" Then
'            sdx3 = sdx3 + Val("" & rconsulta.Fields("total"))
'            End If
'            If Trim("" & rconsulta.Fields("acu")) = "V" Then
'            sdx4 = sdx4 + Val("" & rconsulta.Fields("total"))
'            End If
'            End If
'
'
'            rconsulta.MoveNext
'     Loop
'
'            v = v + 1
'            objExcel.ActiveSheet.Cells(v, h + 6) = "Total"
'            objExcel.ActiveSheet.Cells(v, h + 8) = sdx1
'            objExcel.ActiveSheet.Cells(v, h + 9) = sdx2
'            objExcel.ActiveSheet.Cells(v, h + 10) = sdx1 - sdx2
'            v = v + 1
'            objExcel.ActiveSheet.Cells(v, h + 8) = sdx3
'            objExcel.ActiveSheet.Cells(v, h + 9) = sdx4
'            objExcel.ActiveSheet.Cells(v, h + 9) = sdx3 - sdx4
'
'
'Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
'MsgBox "Proceso Terminado ", 48, "Aviso"
'Exit Sub
'cmd1561212_err:
'MsgBox "Error en exportacion excell " + error$, 48, "Aviso"
'Exit Sub
'
'End Sub

Private Sub dkj8933_Click()

    Dim v, h As Long

    Dim found       As Integer

    Dim I           As Integer

    Dim R           As Long

    Dim sdx         As Double

    Dim sdx1        As Double

    Dim sdx2        As Double

    Dim sdx3        As Double

    Dim sdx4        As Double

    Dim xingreso    As Double

    Dim xegreso     As Double
 
    Dim Heading(12) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd1561212_err

    If MsgBox("Desea Generar reporte ", 1, "Aviso") <> 1 Then Exit Sub
    If rconsulta.RecordCount = 0 Then Exit Sub
    rconsulta.MoveFirst
   
    Heading(1) = "Local"
    Heading(2) = "Tipo"
    Heading(3) = "Serie"
    Heading(4) = "Numero"
    Heading(5) = "Fecha"
    Heading(6) = "Codigo"
    Heading(7) = "Nombre"
    Heading(8) = "M"
    Heading(9) = "Ingreso"
    Heading(10) = "Egreso"
    Heading(11) = "Observación"
   
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_Excelre(12, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
    v = 5

    h = 1
    sdx = 0
    sdx1 = 0
    sdx2 = 0
    sdx3 = 0
    sdx4 = 0

    '    objExcel.ActiveSheet.Cells(v, h + 1) = "Reporte de ingresos Egresos"
    '    v = v + 1
    '    objExcel.ActiveSheet.Cells(v, h + 1) = "FechaI:" & fechai & " Fechaf:" & fechaf
    '    v = v + 1

    objExcel.ActiveSheet.Cells(1, 6) = "     SEGUIMIENTO DE COMPROBANTES"
    objExcel.ActiveSheet.Cells(1, 6).Font.bold = True
    objExcel.ActiveSheet.Cells(1, 6).Font.Size = 14
    objExcel.ActiveSheet.Cells(1, 6).Font.color = RGB(0, 112, 184)
    
    objExcel.ActiveSheet.Cells(2, 1) = "FECHA INICIO  " + fechai
    objExcel.ActiveSheet.Cells(2, 5) = "FECHA FIN  " + fechaf

    Do

        If rconsulta.EOF Then Exit Do
        objExcel.ActiveSheet.Cells(v, h) = "'" & rconsulta.Fields("local")
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & rconsulta.Fields("tipo")
        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & rconsulta.Fields("serie")
        objExcel.ActiveSheet.Cells(v, h + 3) = "'" & rconsulta.Fields("numero")
        objExcel.ActiveSheet.Cells(v, h + 4) = "'" & rconsulta.Fields("fecha")
        objExcel.ActiveSheet.Cells(v, h + 5) = "'" & rconsulta.Fields("codigo")
        objExcel.ActiveSheet.Cells(v, h + 6) = "'" & rconsulta.Fields("nombre")
        objExcel.ActiveSheet.Cells(v, h + 7) = "'" & rconsulta.Fields("moneda")
        xingreso = 0
        xegreso = 0

        If Trim("" & rconsulta.Fields("acu")) = "W" Then
            xingreso = Val("" & rconsulta.Fields("total"))

        End If

        If Trim("" & rconsulta.Fields("acu")) = "V" Then
            xegreso = Val("" & rconsulta.Fields("total"))

        End If

        objExcel.ActiveSheet.Cells(v, h + 8) = xingreso
        objExcel.ActiveSheet.Cells(v, h + 9) = xegreso
        objExcel.ActiveSheet.Cells(v, h + 10) = "'" & rconsulta.Fields("observa")
            
        v = v + 1

        If Trim("" & rconsulta.Fields("moneda")) = "S" Then
            If Trim("" & rconsulta.Fields("acu")) = "W" Then
                sdx1 = sdx1 + Val("" & rconsulta.Fields("total"))

            End If

            If Trim("" & rconsulta.Fields("acu")) = "V" Then
                sdx2 = sdx2 + Val("" & rconsulta.Fields("total"))

            End If

        End If
            
        If Trim("" & rconsulta.Fields("moneda")) = "D" Then
            If Trim("" & rconsulta.Fields("acu")) = "W" Then
                sdx3 = sdx3 + Val("" & rconsulta.Fields("total"))

            End If

            If Trim("" & rconsulta.Fields("acu")) = "V" Then
                sdx4 = sdx4 + Val("" & rconsulta.Fields("total"))

            End If

        End If

        rconsulta.MoveNext
    Loop
                 
    v = v + 1
    objExcel.ActiveSheet.Cells(v, h + 6) = "Total"
    objExcel.ActiveSheet.Cells(v, h + 7) = "S"
    objExcel.ActiveSheet.Cells(v, h + 8) = sdx1
    objExcel.ActiveSheet.Cells(v, h + 9) = sdx2
    objExcel.ActiveSheet.Cells(v, h + 10) = sdx1 - sdx2
            
    Dim k As Integer

    For k = 7 To 11
        objExcel.ActiveSheet.Cells(v, k).Font.bold = True
        objExcel.ActiveSheet.Cells(v, k).Interior.color = RGB(248, 243, 53)
    Next
            
    v = v + 1
    objExcel.ActiveSheet.Cells(v, h + 7) = "D"
    objExcel.ActiveSheet.Cells(v, h + 8) = sdx3
    objExcel.ActiveSheet.Cells(v, h + 9) = sdx4
    objExcel.ActiveSheet.Cells(v, h + 9) = sdx3 - sdx4

    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    MsgBox "Proceso Terminado ", 48, "Aviso"
    Exit Sub
cmd1561212_err:
    MsgBox "Error en exportacion excell " + error$, 48, "Aviso"
    Exit Sub

End Sub

'' 30/11/2017 Mejora reporte ingresos/egresos

Private Sub dnu823_Click()

    Dim found As Integer

    If local1 = "%" Then
        MsgBox "Seleccione Local ", 48, "Aviso"
        Exit Sub

    End If

    found = copiar_recibos()

    If found = 0 Then
        MsgBox "Error al copiar archivo temporal", 24, "Aviso"
        Exit Sub

    End If

    fgusuario = "_l" & gusuario
    found = copiar_tmpfpagoR()

    If found = 0 Then
        MsgBox "No se puede copiar temporal tmpfpagor", 48, "Aviso"
        Exit Sub

    End If

    gofpago = "fpagov"
    fgusuario = "_r" & gusuario
    trecaja.local1 = extra_loquesea(local1)
    trecaja.cajero = gusuario
    trecaja.Caption = explreci.Caption
    trecaja.afecta = afecta
    trecaja.acu = acu
    trecaja.bandera = "NUEVO"
    trecaja.caja = "00"
    trecaja.turno = "1"
    trecaja.Show 1
    sql_recibos

End Sub

Private Sub Form_Activate()

    If estaya = "" Then
        fechai = Format(Now, "dd/mm/yyyy")
        fechaf = Format(Now, "dd/mm/yyyy")
        carga_inicial
        sql_recibos
        estaya = "1"

    End If

End Sub

Sub carga_inicial()

    Dim mytablex As New ADODB.Recordset

    local1.Clear
    local1.AddItem "%"
    mytablex.Open "select * from tlocal ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        local1.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    local1.ListIndex = 0

    If local1.ListCount = 2 Then
        local1.ListIndex = 1

    End If

    cajero.Clear
    cajero.AddItem "%"
    mytablex.Open "select * from vendedor ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        cajero.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    cajero.ListIndex = 0

    tipo.Clear
    tipo.AddItem "%"
    mytablex.Open "select * from tipo ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("tipodoc") = "W" Or "" & mytablex.Fields("tipodoc") = "V" Then
            tipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

        End If

        mytablex.MoveNext
    Loop
    mytablex.Close
    tipo.ListIndex = 0

    caja.Clear
    caja.AddItem "%"
    mytablex.Open "select * from parameca ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("terminal") = "C" Then
            caja.AddItem "" & mytablex.Fields("caja") & "|" & mytablex.Fields("DESCRIPCIO")

        End If

        mytablex.MoveNext
    Loop
    mytablex.Close
    caja.ListIndex = 0

    turno.Clear
    turno.AddItem "%"
    mytablex.Open "select * from turno ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        turno.AddItem "" & mytablex.Fields("turno") & "|" & mytablex.Fields("DESCRIPCIO")
        mytablex.MoveNext
    Loop
    mytablex.Close
    turno.ListIndex = 0

    concepto.Clear
    concepto.AddItem "%"
    mytablex.Open "select * from concepto ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        concepto.AddItem Trim("" & mytablex.Fields("concepto")) & "|" & mytablex.Fields("DESCRIPCIO")
        mytablex.MoveNext
    Loop
    mytablex.Close
    concepto.ListIndex = 0

    subconcepto.Clear
    subconcepto.AddItem "%"
    subconcepto.ListIndex = 0

End Sub

Private Sub Form_Load()
    tacu.Clear
    tacu.AddItem "%"
    tacu.AddItem "W|INGRESO"
    tacu.AddItem "V|EGRESO"
    tacu.ListIndex = 0

    tipoclie.Clear
    tipoclie.AddItem "%"
    tipoclie.AddItem "C"
    tipoclie.AddItem "P"
    tipoclie.AddItem "V"
    tipoclie.ListIndex = 0

End Sub

'Reporte de ingresos (Cobranzas) CONTASIS
Private Sub Label2_Click()

    Dim v, h As Long

    Dim found       As Integer

    Dim I           As Integer

    Dim R           As Long

    Dim sdx         As Double

    Dim sdx1        As Double

    Dim sdx2        As Double

    Dim sdx3        As Double

    Dim sdx4        As Double

    Dim xingreso    As Double

    Dim xegreso     As Double
 
    Dim Heading(20) As String

    On Error GoTo cmd1561212_err

    If MsgBox("Desea Generar reporte ", 1, "Aviso") <> 1 Then Exit Sub
    If rconsulta.RecordCount = 0 Then Exit Sub
    rconsulta.MoveFirst
   
    Heading(1) = "FECHA CANCELACION"
    Heading(2) = "DOCUMENTO"
    Heading(3) = "NUMERO"
    Heading(4) = "CUENTA CONTABLE"
    Heading(5) = "MONEDA"
    Heading(6) = "IMPORTE TOTAL"
    Heading(7) = "T.C"
    
    Heading(8) = "DOCUMENTO"
    Heading(9) = "NUMERO"
    Heading(10) = "FECHA DOCUMENTO"
    Heading(11) = "FECHA VENCIMIENTO"
    
    Heading(12) = "NRO DOC CLIENTE"
    Heading(13) = "APELLIDOS Y NOMBRES,RAZON SOCIAL"
    Heading(14) = "IMPORTE S/"
    Heading(15) = "IMPORTE US$."
    Heading(16) = "CUENTA CONTABLE"
    Heading(17) = "MEDIO DE PAGO"
    Heading(18) = "GLOSA"
    Heading(19) = "CENTRO DE COSTOS 1"
    Heading(20) = "CENTRO DE COSTOS 2"
   
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_ExcelRepCobranzasContasis(20, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
    v = 4

    h = 1
    sdx = 0
    sdx1 = 0
    sdx2 = 0
    sdx3 = 0
    sdx4 = 0

    objExcel.ActiveSheet.Cells(1, 4) = "FORMATO IMPORTACIÓN DE COBRANZAS - SISTEMA EXPERTO CONTABLE 16.00 - Contasis"
    objExcel.ActiveSheet.Cells(1, 4).Font.bold = True
    objExcel.ActiveSheet.Cells(1, 4).Font.Size = 14
    objExcel.ActiveSheet.Cells(1, 4).Font.color = RGB(0, 112, 184)
    
    objExcel.ActiveSheet.Cells(2, 1) = "FECHA INICIO  " + fechai
    objExcel.ActiveSheet.Cells(2, 5) = "FECHA FIN  " + fechaf

    Do

        If rconsulta.EOF Then Exit Do
        objExcel.ActiveSheet.Cells(v, h) = "'" & rconsulta.Fields("fecha")
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & rconsulta.Fields("tipo")
        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & rconsulta.Fields("numero")
        objExcel.ActiveSheet.Cells(v, h + 3) = "'" & busca_CuentasContables(rconsulta.Fields("tipo"), 3)
        objExcel.ActiveSheet.Cells(v, h + 4) = "'" & rconsulta.Fields("moneda")
        objExcel.ActiveSheet.Cells(v, h + 5) = Val("" & rconsulta.Fields("total"))
        objExcel.ActiveSheet.Cells(v, h + 6) = Val("" & rconsulta.Fields("paridad"))
        'objExcel.ActiveSheet.Cells(v, h + 7) = "'" & rconsulta.Fields("moneda")
        
        xingreso = 0
        xegreso = 0
        '            If Trim("" & rconsulta.Fields("acu")) = "W" Then
        '            xingreso = Val("" & rconsulta.Fields("total"))
        '            End If
        '            If Trim("" & rconsulta.Fields("acu")) = "V" Then
        '            xegreso = Val("" & rconsulta.Fields("total"))
        '            End If
        '            objExcel.ActiveSheet.Cells(v, h + 8) = xingreso
        '            objExcel.ActiveSheet.Cells(v, h + 9) = xegreso
        objExcel.ActiveSheet.Cells(v, h + 16) = "CONTADO"
        objExcel.ActiveSheet.Cells(v, h + 17) = "'" & rconsulta.Fields("observa")
            
        v = v + 1
        '            If Trim("" & rconsulta.Fields("moneda")) = "S" Then
        '            If Trim("" & rconsulta.Fields("acu")) = "W" Then
        '            sdx1 = sdx1 + Val("" & rconsulta.Fields("total"))
        '            End If
        '            If Trim("" & rconsulta.Fields("acu")) = "V" Then
        '            sdx2 = sdx2 + Val("" & rconsulta.Fields("total"))
        '            End If
        '            End If
        '
        '            If Trim("" & rconsulta.Fields("moneda")) = "D" Then
        '            If Trim("" & rconsulta.Fields("acu")) = "W" Then
        '            sdx3 = sdx3 + Val("" & rconsulta.Fields("total"))
        '            End If
        '            If Trim("" & rconsulta.Fields("acu")) = "V" Then
        '            sdx4 = sdx4 + Val("" & rconsulta.Fields("total"))
        '            End If
        '            End If
        rconsulta.MoveNext
    Loop
                 
    v = v + 1
    objExcel.ActiveSheet.Cells(v, h + 6) = "Total"
    objExcel.ActiveSheet.Cells(v, h + 7) = "S"
    objExcel.ActiveSheet.Cells(v, h + 8) = sdx1
    objExcel.ActiveSheet.Cells(v, h + 9) = sdx2
    objExcel.ActiveSheet.Cells(v, h + 10) = sdx1 - sdx2
            
    Dim k As Integer

    For k = 7 To 11
        objExcel.ActiveSheet.Cells(v, k).Font.bold = True
        objExcel.ActiveSheet.Cells(v, k).Interior.color = RGB(248, 243, 53)
    Next
            
    v = v + 1
    objExcel.ActiveSheet.Cells(v, h + 7) = "D"
    objExcel.ActiveSheet.Cells(v, h + 8) = sdx3
    objExcel.ActiveSheet.Cells(v, h + 9) = sdx4
    objExcel.ActiveSheet.Cells(v, h + 9) = sdx3 - sdx4

    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    MsgBox "Proceso Terminado ", 48, "Aviso"
    Exit Sub
cmd1561212_err:
    MsgBox "Error en exportacion excell " + error$, 48, "Aviso"
    Exit Sub

End Sub

'Reporte de ingresos (Cobranzas) CONTASIS

Private Sub lfo3434_Click()

    If Frame2.Visible = True Then
        Frame2.Visible = False
        Exit Sub

    End If

    trecitot.Hide
    Unload trecitot

End Sub

Sub sql_recibos()

    On Error GoTo cmd37_err

    Dim buf As String

    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub
    buf = "select * from recibo where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If local1 <> "%" Then
        buf = buf & " and local='" & extra_loquesea(local1) & "'"

    End If

    If codigo <> "%" Then
        buf = buf & " and codigo like '" & codigo & "'"

    End If

    If nombre <> "%" Then
        buf = buf & " and nombre like '" & nombre & "'"

    End If

    If concepto <> "%" Then
        buf = buf & " and concepto='" & extra_loquesea(concepto) & "'"

    End If

    If subconcepto <> "%" Then
        buf = buf & " and subconcepto='" & extra_loquesea(subconcepto) & "'"

    End If

    '' 30/11/2017 Mejora reporte ingresos/egresos
    buf = buf & " and estado='2' "
    '' 30/11/2017 Mejora reporte ingresos/egresos

    If cajero <> "%" Then
        buf = buf & " and usuario='" & extra_loquesea(cajero) & "'"

    End If

    If tipo <> "%" Then
        buf = buf & " and tipo='" & extra_loquesea(tipo) & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja='" & extra_loquesea(caja) & "'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno='" & extra_loquesea(turno) & "'"

    End If

    'buf = buf & " and acu='" & acu & "'"

    'Reporte de ingresos (Cobranzas) CONTASIS
    'buf = buf & " order by fecha,"
    buf = buf & " order by tipo,fecha, str(numero)"
    'Reporte de ingresos (Cobranzas) CONTASIS

    'MsgBox buf
   
    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open buf, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True And rconsulta.BOF = True Then
   
    End If

    'MsgBox ""
    Set DBGrid2.DataSource = rconsulta
  
    'MsgBox ""
   
    sumar_recibos rconsulta

    If rconsulta.RecordCount > 0 Then

        'dbgrid2.SetFocus
    End If

    Exit Sub
cmd37_err:
    MsgBox "Error en select " & error$, 48, "Aviso"
    Exit Sub

End Sub

Sub sumar_recibos(mytablex As ADODB.Recordset)

    Dim xisoles   As Double

    Dim xidolares As Double

    Dim xesoles   As Double

    Dim xedolares As Double

    Dim xsoles    As Double

    Dim xdolares  As Double

    Dim sdx1      As Double

    Dim sdx       As Double

    On Error GoTo cmd345_err

    xisoles = 0
    xidolares = 0
    xesoles = 0
    xedolares = 0

    xsoles = 0
    xdolares = 0

    soles = ""
    dolares = ""
    isoles = ""
    idolares = ""
    esoles = ""
    edolares = ""

    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("acu") = "W" Then
            If "" & mytablex.Fields("moneda") = "S" Then
                xisoles = xisoles + Val("" & mytablex.Fields("total"))

            End If

            If "" & mytablex.Fields("moneda") = "D" Then
                xidolares = xidolares + Val("" & mytablex.Fields("total"))

            End If

        End If

        If "" & mytablex.Fields("acu") = "V" Then
            If "" & mytablex.Fields("moneda") = "S" Then
                xesoles = xesoles + Val("" & mytablex.Fields("total"))

            End If

            If "" & mytablex.Fields("moneda") = "D" Then
                xedolares = xedolares + Val("" & mytablex.Fields("total"))

            End If

        End If

        mytablex.MoveNext
    Loop
    isoles = Format(xisoles, "0.00")
    idolares = Format(xidolares, "0.00")

    esoles = Format(xesoles, "0.00")
    edolares = Format(xedolares, "0.00")

    sdx = xisoles - xesoles
    sdx1 = xidolares - xedolares

    soles = Format(sdx, "0.00")
    dolares = Format(sdx1, "0.00")

    Exit Sub
cmd345_err:
    MsgBox "Aviso en sumar recibos " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub proceso_impresion1(bxtipoclie As String, _
                       bxlocal As String, _
                       bxtipo As String, _
                       bxserie As String, _
                       bxnumero As String, _
                       xvacu As String)

    Dim found    As Integer

    Dim archivot As String

    On Error GoTo cmd6_err:

    cerrar_archivo
    factura_formatox bxtipoclie, bxlocal, "" & bxtipo, "" & bxserie, "" & bxnumero, "", xvacu
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    Exit Sub
cmd6_err:
    MsgBox "Mensaje, Error al iniciar Impresion " & error$
    Exit Sub

End Sub

Sub factura_formatox(tipoclie As String, _
                     bxlocal As String, _
                     bxtipo As String, _
                     bxserie As String, _
                     bxnumero As String, _
                     ascopia As String, _
                     xvacu As String)

    Dim vacu            As String

    Dim mytablex        As New ADODB.Recordset

    Dim found           As Integer

    Dim nro_lineas      As Integer

    Dim contando        As Integer

    Dim faltante        As Integer

    Dim I               As Integer

    Dim archivo_formato As String

    Dim xtipoarchivo    As String

    Dim mytabley        As New ADODB.Recordset

    On Error GoTo cmd450009_err

    If tipoclie = "C" Then
        xtipoarchivo = "CUENTACD"

    End If

    If tipoclie = "P" Then
        xtipoarchivo = "CUENTAPD"

    End If

    If tipoclie = "V" Then
        xtipoarchivo = "CUENTACD"

    End If

    vacu = ""
    nro_lineas = 13
    contando = 0
    FileName = globaldir & "\temporal\" & gusuario & ".txt"
       
    found = borra_nombre("" & FileName)
       
    archivo_formato = busca_archivo_formato(bxtipo)

    If Len(archivo_formato) = 0 Then
        MsgBox "No existe archivo formato ", 48, "Aviso"
        Exit Sub

    End If

    'recibo
       
    mytabley.Open "select * from recibo where local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount = 0 Then
        mytabley.Close
        Exit Sub

    End If
       
    'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    'found = proceso_formatos(archivo_formato, mytabley, "{", "}", "RECIBO", "RECIBO", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    found = proceso_formatos(archivo_formato, mytabley, "{", "}", "RECIBO", "RECIBO", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    vacu = "" & mytabley.Fields("acu")
       
    'mytabley.Close
    '
    'detalle
    flag_contando = 0
    'Set mydbx = OpenDatabase(globaldat, False, False, "foxpro 2.5;")
    'MsgBox xtipoarchivo
       
    mytablex.Open "select * from " & xtipoarchivo & " where local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        Do

            If mytablex.EOF Then Exit Do
            flag_contando = contando + 1
            'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
            'found = proceso_formatos(archivo_formato, mytabley, "{", "}", "RECIBO", "RECIBO", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
            found = proceso_formatos(archivo_formato, mytabley, "{", "}", "RECIBO", "RECIBO", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
            'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
          
            contando = contando + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
    '
    'If bxtipo <> "1" And bxtipo <> "2" And bxtipo <> "5" Then
    '   If contando < nro_lineas Then
    '      For i = contando To nro_lineas
    '          Open filename For Append As #1
    '          found = formateaa("", 1, 2, 0)
    '          Close #1
    '      Next i
    '   End If
    'End If
    '----- SUBTOTAL
       
    'Set mydbx = OpenDatabase(globaldat, False, False, "foxpro 2.5;")
    'Set mytablex = mydbxglo.OpenTable("RECIBO")
    'mytablex.Index = "RECIBO"
       
    'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    'found = proceso_formatos(archivo_formato, mytabley, "{", "}", "RECIBO", "RECIBO", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    found = proceso_formatos(archivo_formato, mytabley, "{", "}", "RECIBO", "RECIBO", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
          
    'mytablex.Close
    '
    'forma de pago
    'Set mydbx = OpenDatabase(globaldat, False, False, "foxpro 2.5;")
        
    mytablex.Open "select * from fpagov where local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        
        Do

            If mytablex.EOF Then Exit Do
            'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
            'found = proceso_formatos(archivo_formato, mytablex, "<", ">", "FPAGOV", "fpagov", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
            found = proceso_formatos(archivo_formato, mytablex, "<", ">", "FPAGOV", "fpagov", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
            'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
    '
    '----------pie de paginatotal  xxxx
    'Set mydbx = OpenDatabase(globaldat, False, False, "foxpro 2.5;")
    'Set mytablex = mydbxglo.OpenTable("recibo")
    'mytablex.Index = "recibo"
       
    'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    'found = proceso_formatos(archivo_formato, mytabley, "^", "&", "RECIBO", "RECIBO", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    found = proceso_formatos(archivo_formato, mytabley, "^", "&", "RECIBO", "RECIBO", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
          
    mytabley.Close
    Exit Sub
cmd450009_err:
    MsgBox "Aviso Proceso en Compilando " & error$, 24, "Aviso"
    Exit Sub

End Sub

Function control_impresionxx(bxtipo As String)

    Dim found As Integer

    Dim sFile As String

    sFile = globaldir & "\temporal\" & gusuario & ".txt"
    found = Imprime_archivojj(sFile, 0, "8", "", "S", "")
    Exit Function
cmd67111_err:

End Function

Function busca_archivo_formato(bxtipo As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from tipo where tipo='" & bxtipo & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_archivo_formato = "" & mytablex.Fields("archivo")

    End If

    mytablex.Close
 
End Function

Public Function Formato_Excelre(Num_Campos As Integer, _
                                Nombre_Campos() As String) As Boolean

    With objExcel.ActiveSheet
        'MsgBox Num_Campos
        'Formato de las celdas de los titulos
        .Range(.Cells(3, 1), .Cells(3, 11)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 11)).Font.bold = True
        .Range(.Cells(3, 1), .Cells(3, 11)).Interior.color = RGB(192, 192, 250)
        
        For I = 1 To Num_Campos Step 1
            .Cells(3, I) = Nombre_Campos(I)
        Next I

        'hasta aki pa colocar los titulos
        
        'a partir de aki ta claro que es pa darle el ancho a las celdas ;-)
       
        .columns("A").ColumnWidth = 5
        .columns("B").ColumnWidth = 5
        .columns("C").ColumnWidth = 5
        .columns("D").ColumnWidth = 12
        .columns("E").ColumnWidth = 10
        .columns("F").ColumnWidth = 15
        .columns("G").ColumnWidth = 20
        .columns("H").ColumnWidth = 3
        .columns("i").ColumnWidth = 10
        .columns("j").ColumnWidth = 10
        .columns("k").ColumnWidth = 12
    
    End With

End Function

Sub carga_subconcepto(buf As String)

    Dim mytablex As New ADODB.Recordset

    subconcepto.Clear
    subconcepto.AddItem "%"
    mytablex.Open "select * from subconcepto where concepto='" & buf & "'", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        subconcepto.AddItem Trim("" & mytablex.Fields("subconcepto")) & "|" & mytablex.Fields("DESCRIPCIO")
        mytablex.MoveNext
    Loop
    mytablex.Close
    subconcepto.ListIndex = 0

End Sub

'Reporte de ingresos (Cobranzas) CONTASIS
Public Function Formato_ExcelRepCobranzasContasis(Num_Campos As Integer, _
                                                  Nombre_Campos() As String) As Boolean

    With objExcel.ActiveSheet
        'Formato de las celdas de los titulos
        .Range(.Cells(3, 1), .Cells(3, 7)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 7)).Font.bold = True
        .Range(.Cells(3, 1), .Cells(3, 7)).Interior.color = RGB(192, 192, 250)
        
        For I = 1 To Num_Campos Step 1
            .Cells(3, I) = Nombre_Campos(I)
 
            .columns("A").ColumnWidth = 20
            .columns("B").ColumnWidth = 13
            .columns("C").ColumnWidth = 13
            .columns("D").ColumnWidth = 18
            .columns("E").ColumnWidth = 10
            .columns("F").ColumnWidth = 15
            .columns("G").ColumnWidth = 8
            .columns("H").ColumnWidth = 15
            .columns("I").ColumnWidth = 15
            .columns("J").ColumnWidth = 20
            .columns("K").ColumnWidth = 20
            .columns("L").ColumnWidth = 20
            .columns("M").ColumnWidth = 50
            .columns("N").ColumnWidth = 15
            .columns("O").ColumnWidth = 15
            .columns("P").ColumnWidth = 17
            .columns("Q").ColumnWidth = 18
            .columns("R").ColumnWidth = 25
            .columns("S").ColumnWidth = 20
            .columns("T").ColumnWidth = 20
   
        Next

    End With
    
    With objExcel.ActiveSheet
        .Range(.Cells(3, 1), .Cells(3, 7)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 7)).Font.bold = True
        .Range(.Cells(3, 8), .Cells(3, 20)).Interior.color = RGB(192, 200, 200)

    End With
    
End Function

'Reporte de ingresos (Cobranzas) CONTASIS

'Reporte de ingresos (Cobranzas) CONTASIS
Function busca_CuentasContables(tipodoc As String, tipocuenta As String) As String

    Dim mytabley As New ADODB.Recordset

    mytabley.Open "SELECT cuenta1,cuenta2,cuenta3  FROM tipo where tipo='" & "" & tipodoc & "'", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount > 0 Then
   
        If tipocuenta = "1" Then        'SubTotal
            busca_CuentasContables = "" & mytabley.Fields("cuenta1")
        ElseIf tipocuenta = "2" Then    'Impuesto
            busca_CuentasContables = "" & mytabley.Fields("cuenta2")
        ElseIf tipocuenta = "3" Then    'Total
            busca_CuentasContables = "" & mytabley.Fields("cuenta3")

        End If

    End If

    '------------------------------------- ------------
    mytabley.Close
 
End Function

'Reporte de ingresos (Cobranzas) CONTASIS
