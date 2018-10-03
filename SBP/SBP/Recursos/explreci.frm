VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form explreci 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibos de Caja"
   ClientHeight    =   9885
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9885
   ScaleWidth      =   15225
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consulta"
      Height          =   5055
      Left            =   3240
      TabIndex        =   3
      Top             =   3360
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
         Picture         =   "explreci.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "explreci.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "explreci.frx":0F5C
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.ComboBox concepto 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   11160
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox subconcepto 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   11160
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox local1 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   11160
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   0
         Width           =   1815
      End
      Begin VB.ComboBox tipoclie 
         BackColor       =   &H00FFFFFF&
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
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   0
         Width           =   615
      End
      Begin VB.ComboBox turno 
         BackColor       =   &H00FFFFFF&
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
         Left            =   8640
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox tipo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   720
         Width           =   1815
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
         Height          =   375
         Left            =   14640
         MaskColor       =   &H00FFFFFF&
         Picture         =   "explreci.frx":216E
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   600
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cajero 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8640
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox codigo 
         Height          =   375
         Left            =   6840
         MaxLength       =   13
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
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   14
         Top             =   0
         Width           =   1815
      End
      Begin VB.TextBox fechaf 
         BackColor       =   &H0080FFFF&
         Height          =   375
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   13
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox caja 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8640
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   0
         Width           =   1815
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
         Picture         =   "explreci.frx":291C
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   2160
         MaskColor       =   &H00E0E0E0&
         Picture         =   "explreci.frx":3B2E
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
         Left            =   1440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "explreci.frx":4D40
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprimir"
         Top             =   0
         Width           =   735
      End
      Begin ChamaleonButton.ChameleonBtn Label19 
         Height          =   810
         Left            =   13080
         TabIndex        =   40
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
         MICON           =   "explreci.frx":5F52
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Concepto"
         Height          =   375
         Left            =   10560
         TabIndex        =   39
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Subconcepto"
         Height          =   375
         Left            =   10560
         TabIndex        =   38
         Top             =   720
         Width           =   615
      End
      Begin VB.Label xcuentaco1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   13560
         TabIndex        =   35
         Top             =   480
         Width           =   105
      End
      Begin VB.Label xcuentaco 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   13680
         TabIndex        =   34
         Top             =   480
         Width           =   105
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
         Height          =   375
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
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         Height          =   375
         Left            =   3000
         TabIndex        =   26
         Top             =   720
         Width           =   1095
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
         Left            =   3000
         TabIndex        =   16
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   375
         Left            =   3000
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
      Height          =   7215
      Left            =   0
      TabIndex        =   32
      Top             =   1200
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   12726
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
      ColumnCount     =   21
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
         DataField       =   "OBSERVA"
         Caption         =   "Observación"
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
         DataField       =   "Concepto"
         Caption         =   "Concepto"
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
         DataField       =   "Subconcepto"
         Caption         =   "Subconcepto"
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
      BeginProperty Column19 
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
      BeginProperty Column20 
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   2
         BeginProperty Column00 
            ColumnWidth     =   225.071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   464.882
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   629.858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1395.213
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
            ColumnWidth     =   2865.26
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   225.071
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   390.047
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   195.024
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   2294.929
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column18 
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   1080
         EndProperty
      EndProperty
   End
   Begin VB.Label dolaresanu 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   49
      Top             =   9240
      Width           =   1575
   End
   Begin VB.Label solesanu 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   48
      Top             =   9240
      Width           =   1695
   End
   Begin VB.Label Label20 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Anulados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   47
      Top             =   9240
      Width           =   1095
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      Height          =   375
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   45
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   44
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dolares"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   43
      Top             =   8520
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Soles"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   42
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Label Label13 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   41
      Top             =   8880
      Width           =   1095
   End
   Begin VB.Label estaya 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3840
      TabIndex        =   24
      Top             =   8280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label dolares 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   10
      Top             =   8880
      Width           =   1575
   End
   Begin VB.Label soles 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   9
      Top             =   8880
      Width           =   1695
   End
   Begin VB.Label afecta 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   8160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label acu 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Menu dnu823 
      Caption         =   "&Nuevo"
   End
   Begin VB.Menu anul923 
      Caption         =   "&Anula"
   End
   Begin VB.Menu dbo912 
      Caption         =   "&Borra"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu Reporte 
      Caption         =   "&Reporte"
   End
   Begin VB.Menu dfkl8823 
      Caption         =   "&Copia"
   End
   Begin VB.Menu lfo3434 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "explreci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rconsulta As New ADODB.Recordset

Private Sub anul923_Click()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd34_err

    ''' 30/11/2017 Correción  General del Sistema Parte I
    'If tipoclie = "%" Then
    '   MsgBox "Seleccione tipo de C.Cliente P.Proveedor V.Vendedor", 48, "Aviso"
    '   Exit Sub
    'End If
    ''' 30/11/2017 Correción  General del Sistema Parte I

    If DBGrid2.columns("e") <> "2" Then
        MsgBox "Debe estar en estado 2", 1, "Aviso"
        Exit Sub

    End If

    If MsgBox("Desea ANULAR el registro " & DBGrid2.columns("numero"), 1, "Aviso") <> "1" Then Exit Sub
    If "" & DBGrid2.columns("acu") = "W" Or "" & DBGrid2.columns("acu") = "V" Then  'ingreso/egreso
        found = descarga_cuentac("" & DBGrid2.columns("local"), "" & DBGrid2.columns("tipo"), "" & DBGrid2.columns("serie"), "" & DBGrid2.columns("numero"), "-1")

    End If

    'found = borra_letra("" & dbgrid2.columns("local"), "" & dbgrid2.columns("tipo"), "" & dbgrid2.columns("serie"), "" & dbgrid2.columns("numero"))
    found = borra_fpagov("" & DBGrid2.columns("local"), "" & DBGrid2.columns("tipo"), "" & DBGrid2.columns("serie"), "" & DBGrid2.columns("numero"))
    rconsulta.Fields("estado") = "1"
    rconsulta.Update
    'cn.Execute ("delete from recibo where local='" & dbgrid2.columns("local") & "' and tipo='" & "" & dbgrid2.columns("tipo") & "' and serie='" & "" & dbgrid2.columns("serie") & "' and numero='" & "" & dbgrid2.columns("numero") & "'")
    sql_recibos
    MsgBox "Proceso Borrado ", 48, "Aviso"
    Exit Sub
cmd34_err:
    MsgBox "Aviso en borra " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub borra923_Click()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd34_err

    If tipoclie = "%" Then
        MsgBox "Seleccione tipo de C.Cliente P.Proveedor V.Vendedor", 48, "Aviso"
        Exit Sub

    End If

    If DBGrid2.columns("e") <> "2" Then
        MsgBox "Debe estar en estado 2", 1, "Aviso"
        Exit Sub

    End If

    If MsgBox("Desea ANULAR el registro " & DBGrid2.columns("numero"), 1, "Aviso") <> "1" Then Exit Sub
    If "" & DBGrid2.columns("acu") = "W" Or "" & DBGrid2.columns("acu") = "V" Then  'ingreso/egreso
        found = descarga_cuentac("" & DBGrid2.columns("local"), "" & DBGrid2.columns("tipo"), "" & DBGrid2.columns("serie"), "" & DBGrid2.columns("numero"), "-1")

    End If

    'found = borra_letra("" & dbgrid2.columns("local"), "" & dbgrid2.columns("tipo"), "" & dbgrid2.columns("serie"), "" & dbgrid2.columns("numero"))
    found = borra_fpagov("" & DBGrid2.columns("local"), "" & DBGrid2.columns("tipo"), "" & DBGrid2.columns("serie"), "" & DBGrid2.columns("numero"))
    rconsulta.Fields("estado") = "1"
    rconsulta.Update
    'cn.Execute ("delete from recibo where local='" & dbgrid2.columns("local") & "' and tipo='" & "" & dbgrid2.columns("tipo") & "' and serie='" & "" & dbgrid2.columns("serie") & "' and numero='" & "" & dbgrid2.columns("numero") & "'")
    sql_recibos
    MsgBox "Proceso Borrado ", 48, "Aviso"
    Exit Sub
cmd34_err:
    MsgBox "Aviso en borra " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub cmdAddEntry_Click()

    Dim found As Integer

    If local1 = "%" Then
        MsgBox "Seleccione Local ", 48, "Aviso"
        Exit Sub

    End If

    If tipoclie = "%" Then
        MsgBox "Seleccione tipo de C.Cliente P.Proveedor V.Vendedor", 48, "Aviso"
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
    trecaja.pagocash.Visible = True
    trecaja.pagocash.Value = 1

    trecaja.xcuentaco = xcuentaco
    trecaja.XCUENTACO1 = XCUENTACO1

    trecaja.local1 = extra_loquesea(local1)
    trecaja.cajero = gusuario
    trecaja.Caption = explreci.Caption
    trecaja.afecta = afecta
    trecaja.acu = acu
    trecaja.tipoclie = tipoclie
    trecaja.tipoclie.Enabled = False
    trecaja.bandera = "NUEVO"
    trecaja.caja = "00"
    trecaja.turno = "1"
    trecaja.Show 1
    sql_recibos

End Sub

Private Sub cmdCancelar_Click()
    lfo3434_Click

End Sub

Private Sub cmdDelete_Click()
    dbo912_Click

End Sub

Private Sub cmdExit_Click()

    If Frame2.Visible = True Then
        Frame2.Visible = False
        Exit Sub

    End If

    explreci.Hide
    Unload explreci

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

    Select Case xcuentaco

        Case "CUENTAC"
            buf1 = "CUENTACD"

        Case "CUENTAP"
            buf1 = "CUENTAPD"

    End Select

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

''' 30/11/2017 Mejora reporte ingresos/egresos

'Private Sub cmdPrint_Click()
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

Private Sub cmdPrint_Click()

    Dim v, h As Long

    Dim found       As Integer

    Dim I           As Integer

    Dim R           As Long

    Dim sdx         As Double
 
    Dim sdx1        As Double

    Dim sdx2        As Double

    Dim sdx3        As Double

    Dim sdx4        As Double
 
    Dim sdx1anu     As Double

    Dim sdx2anu     As Double

    Dim sdx3anu     As Double

    Dim sdx4anu     As Double
 
    Dim xingreso    As Double

    Dim xegreso     As Double
 
    Dim xingresoanu As Double

    Dim xegresoanu  As Double
    
    Dim Heading(13) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd1561212_err

    If MsgBox("Desea Generar reporte ", 1, "Aviso") <> 1 Then Exit Sub
    If rconsulta.RecordCount = 0 Then Exit Sub
    rconsulta.MoveFirst
   
    Heading(1) = "Estado"
    Heading(2) = "Local"
    Heading(3) = "Tipo"
    Heading(4) = "Serie"
    Heading(5) = "Numero"
    Heading(6) = "Fecha"
    Heading(7) = "Codigo"
    Heading(8) = "Nombre"
    Heading(9) = "M"
    Heading(10) = "Ingreso"
    Heading(11) = "Egreso"
    Heading(12) = "Observación"
   
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_Excelre(13, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
    v = 5
    h = 1
    sdx = 0
    sdx1 = 0
    sdx2 = 0
    sdx3 = 0
    sdx4 = 0

    sdxanu = 0
    sdx1anu = 0
    sdx2anu = 0
    sdx3anu = 0
    sdx4anu = 0

    If Me.Caption = "INGRESO DINERO" Then
        objExcel.ActiveSheet.Cells(1, 6) = "     REPORTE DE INGRESO DE DINEROS"
    ElseIf Me.Caption = "EGRESO DINERO" Then
        objExcel.ActiveSheet.Cells(1, 6) = "     REPORTE DE EGRESOS DE DINERO"

    End If
  
    objExcel.ActiveSheet.Cells(1, 6).Font.bold = True
    objExcel.ActiveSheet.Cells(1, 6).Font.Size = 14
    objExcel.ActiveSheet.Cells(1, 6).Font.color = RGB(0, 112, 184)
    
    objExcel.ActiveSheet.Cells(2, 1) = "FECHA INICIO  " + fechai
    objExcel.ActiveSheet.Cells(2, 5) = "FECHA FIN  " + fechaf
    
    'v = v + 1
    ' objExcel.ActiveSheet.Cells(v, h + 1) = "FechaI:" & fechai & " Fechaf:" & fechaf
    'v = v + 1

    Do

        If rconsulta.EOF Then Exit Do
            
        objExcel.ActiveSheet.Cells(v, h) = "'" & rconsulta.Fields("estado")
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & rconsulta.Fields("local")
        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & rconsulta.Fields("tipo")
        objExcel.ActiveSheet.Cells(v, h + 3) = "'" & rconsulta.Fields("serie")
        objExcel.ActiveSheet.Cells(v, h + 4) = "'" & rconsulta.Fields("numero")
        objExcel.ActiveSheet.Cells(v, h + 5) = "'" & rconsulta.Fields("fecha")
        objExcel.ActiveSheet.Cells(v, h + 6) = "'" & rconsulta.Fields("codigo")
        objExcel.ActiveSheet.Cells(v, h + 7) = "'" & rconsulta.Fields("nombre")
        objExcel.ActiveSheet.Cells(v, h + 8) = "'" & rconsulta.Fields("moneda")
            
        xingreso = 0
        xegreso = 0
        xingresoanu = 0
        xegresoanu = 0
            
        If Trim("" & rconsulta.Fields("acu")) = "W" Then
            xingreso = Val("" & rconsulta.Fields("total"))

        End If
            
        If Trim("" & rconsulta.Fields("acu")) = "V" Then
            xegreso = Val("" & rconsulta.Fields("total"))

        End If
            
        objExcel.ActiveSheet.Cells(v, h + 9) = xingreso
        objExcel.ActiveSheet.Cells(v, h + 10) = xegreso
        objExcel.ActiveSheet.Cells(v, h + 11) = "'" & rconsulta.Fields("observa")
            
        v = v + 1
            
        If Trim("" & rconsulta.Fields("moneda")) = "S" Then
                
            If Trim("" & rconsulta.Fields("acu")) = "W" And "" & rconsulta.Fields("estado") = "2" Then
                sdx1 = sdx1 + Val("" & rconsulta.Fields("total"))

            End If

            If Trim("" & rconsulta.Fields("acu")) = "W" And "" & rconsulta.Fields("estado") = "1" Then
                sdx1anu = sdx1anu + Val("" & rconsulta.Fields("total"))

            End If
   
            If Trim("" & rconsulta.Fields("acu")) = "V" And "" & rconsulta.Fields("estado") = "2" Then
                sdx2 = sdx2 + Val("" & rconsulta.Fields("total"))

            End If

            If Trim("" & rconsulta.Fields("acu")) = "V" And "" & rconsulta.Fields("estado") = "1" Then
                sdx2anu = sdx2anu + Val("" & rconsulta.Fields("total"))

            End If
            
        End If
            
        If Trim("" & rconsulta.Fields("moneda")) = "D" Then
                
            If Trim("" & rconsulta.Fields("acu")) = "W" And "" & rconsulta.Fields("estado") = "2" Then
                sdx3 = sdx3 + Val("" & rconsulta.Fields("total"))

            End If

            If Trim("" & rconsulta.Fields("acu")) = "W" And "" & rconsulta.Fields("estado") = "1" Then
                sdx3anu = sdx3anu + Val("" & rconsulta.Fields("total"))

            End If
                
            If Trim("" & rconsulta.Fields("acu")) = "V" And "" & rconsulta.Fields("estado") = "2" Then
                sdx4 = sdx4 + Val("" & rconsulta.Fields("total"))

            End If

            If Trim("" & rconsulta.Fields("acu")) = "V" And "" & rconsulta.Fields("estado") = "1" Then
                sdx4anu = sdx4anu + Val("" & rconsulta.Fields("total"))

            End If
            
        End If
            
        rconsulta.MoveNext
    Loop
     
    v = v + 1
            
    objExcel.ActiveSheet.Cells(v, h + 7) = "Total"
            
    objExcel.ActiveSheet.Cells(v, h + 8) = "S"
    objExcel.ActiveSheet.Cells(v + 1, h + 8) = "D"
            
    objExcel.ActiveSheet.Cells(v + 3, h + 7) = "Anulados"
    objExcel.ActiveSheet.Cells(v + 3, h + 7).Font.bold = True
            
    objExcel.ActiveSheet.Cells(v + 3, h + 8) = "S"
    objExcel.ActiveSheet.Cells(v + 4, h + 8) = "D"
            
    objExcel.ActiveSheet.Cells(v, h + 9) = sdx1 'total soles ingresos
    objExcel.ActiveSheet.Cells(v, h + 10) = sdx2 ' total soles egresos
            
    objExcel.ActiveSheet.Cells(v + 3, h + 9) = sdx1anu 'anulados soles ingresos
    objExcel.ActiveSheet.Cells(v + 3, h + 10) = sdx2anu ' anulados soles egresos
            
    Dim k As Integer

    For k = 8 To 11
        objExcel.ActiveSheet.Cells(v, k).Font.bold = True
        objExcel.ActiveSheet.Cells(v, k).Interior.color = RGB(248, 243, 53)
    Next

    ''
    'objExcel.ActiveSheet.Cells(v, h + 11) = sdx1 - sdx2
    objExcel.ActiveSheet.Cells(v, h + 11) = ""
    ''
            
    v = v + 1
            
    objExcel.ActiveSheet.Cells(v, h + 9) = sdx3 'total dolares ingresos
    objExcel.ActiveSheet.Cells(v, h + 10) = sdx4 ' total dolares egresos
                    
    objExcel.ActiveSheet.Cells(v + 3, h + 9) = sdx3anu 'anulados dolares ingresos
    objExcel.ActiveSheet.Cells(v + 3, h + 10) = sdx4anu ' anulados dolares egresos
            
    objExcel.ActiveSheet.Cells(v, h + 10) = sdx3 - sdx4

    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    MsgBox "Proceso Terminado ", 48, "Aviso"
    Exit Sub
cmd1561212_err:
    MsgBox "Error en exportacion excell " + error$, 48, "Aviso"
    Exit Sub

End Sub

''' 30/11/2017 Mejora reporte ingresos/egresos

Private Sub cmdSort_Click()

End Sub

Private Sub Command1_Click()
    'If tipoclie = "%" Then
    '   MsgBox "Seleccione tipo de C.Cliente P.Proveedor V.Vendedor", 48, "Aviso"
    '   Exit Sub
    'End If

    sql_recibos

End Sub

Private Sub concepto_Click()

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    If concepto = "%" Then Exit Sub
    buf = Trim("" & extra_loquesea(concepto))
    carga_subconcepto "" & buf

End Sub

Private Sub dbo912_Click()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd34_err

    '' 30/11/2017 Mejora reporte ingresos/egresos
    'If tipoclie = "%" Then
    '   MsgBox "Seleccione tipo de C.Cliente P.Proveedor V.Vendedor", 48, "Aviso"
    '   Exit Sub
    'End If
    '' 30/11/2017 Mejora reporte ingresos/egresos

    If DBGrid2.columns("e") <> "0" Then
        MsgBox "Debe estar en estado 0", 1, "Aviso"
        Exit Sub

    End If

    If MsgBox("Desea Borra el registro " & DBGrid2.columns("numero"), 1, "Aviso") <> "1" Then Exit Sub
    If "" & DBGrid2.columns("acu") = "W" Or "" & DBGrid2.columns("acu") = "V" Then  'ingreso/egreso
        found = descarga_cuentac("" & DBGrid2.columns("local"), "" & DBGrid2.columns("tipo"), "" & DBGrid2.columns("serie"), "" & DBGrid2.columns("numero"), "-1")

    End If

    'found = borra_letra("" & dbgrid2.columns("local"), "" & dbgrid2.columns("tipo"), "" & dbgrid2.columns("serie"), "" & dbgrid2.columns("numero"))
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

Function borra_letra(xlocal As String, _
                     xtipo As String, _
                     xserie As String, _
                     xnumero As String)
    cn.Execute ("delete from letrav where local='" & xlocal & "' and tipor='" & xtipo & "' and serier='" & xserie & "' and numeror='" & xnumero & "'")

End Function

Private Sub dfkl8823_Click()

    On Error GoTo cdm99_err
   
    ''' 30/11/2017 Correción  General del Sistema Parte I
   
    'If tipoclie = "%" Then
    '   MsgBox "Seleccione tipo de C.Cliente P.Proveedor V.Vendedor", 48, "Aviso"
    '   Exit Sub
    'End If
    ''' 30/11/2017 Correción  General del Sistema Parte I

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

Private Sub dnu823_Click()

    Dim found As Integer

    If local1 = "%" Then
        MsgBox "Seleccione Local ", 48, "Aviso"
        Exit Sub

    End If

    If tipoclie = "%" Then
        MsgBox "Seleccione tipo de C.Cliente P.Proveedor V.Vendedor", 48, "Aviso"
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
    trecaja.pagocash.Visible = True
    trecaja.pagocash.Value = 1

    trecaja.xcuentaco = xcuentaco
    trecaja.XCUENTACO1 = XCUENTACO1

    trecaja.local1 = extra_loquesea(local1)
    trecaja.cajero = gusuario
    trecaja.Caption = explreci.Caption
    trecaja.afecta = afecta
    trecaja.acu = acu
    trecaja.tipoclie = tipoclie
    trecaja.tipoclie.Enabled = False
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

        tipoclie.Clear
        tipoclie.AddItem "%"

        If acu = "W" Then
            tipoclie.AddItem "C"
            tipoclie.AddItem "V"

        End If

        If acu = "V" Then
            tipoclie.AddItem "P"
            tipoclie.AddItem "C"
            tipoclie.AddItem "V"

        End If

        tipoclie.ListIndex = 0

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

        'If "" & mytablex.Fields("tipodoc") = "W" Or "" & mytablex.Fields("tipodoc") = "V" Then
        If "" & mytablex.Fields("tipodoc") = acu Then
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
    subconcepto.Clear

    concepto.AddItem "%"
    mytablex.Open "select * from concepto ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        concepto.AddItem "" & mytablex.Fields("concepto") & "|" & mytablex.Fields("DESCRIPCIO")
        mytablex.MoveNext
    Loop
    mytablex.Close
    concepto.ListIndex = 0

    subconcepto.AddItem "%"
    subconcepto.ListIndex = 0

End Sub

Private Sub Label26_Click()
    sql_recibos

End Sub

Private Sub Label19_Click()

    sql_recibos

End Sub

Private Sub lfo3434_Click()

    If Frame2.Visible = True Then
        Frame2.Visible = False
        Exit Sub

    End If

    explreci.Hide
    Unload explreci

End Sub

Sub sql_recibos()

    On Error GoTo cmd37_err

    Dim buf As String

    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub
    buf = "select * from recibo where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If concepto <> "%" Then
        buf = buf & " and concepto='" & extra_loquesea(concepto) & "'"

    End If

    If subconcepto <> "%" Then
        buf = buf & " and subconcepto='" & extra_loquesea(subconcepto) & "'"

    End If

    If local1 <> "%" Then
        buf = buf & " and local='" & extra_loquesea(local1) & "'"

    End If

    If codigo <> "%" Then
        buf = buf & " and codigo like '" & codigo & "'"

    End If

    If nombre <> "%" Then
        buf = buf & " and nombre like '" & nombre & "'"

    End If

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

    buf = buf & " and acu='" & acu & "'"
    buf = buf & " order by fecha"
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

'' 30/11/2017 Correción  General del Sistema Parte I. Correcion suma de recibos ene stado 2 y 1

'Sub sumar_recibos(mytablex As ADODB.Recordset)
'Dim xsoles As Double
'Dim xdolares As Double
'On Error GoTo cmd345_err
'xsoles = 0
'xdolares = 0
'soles = "0.00"
'dolares = "0.00"
'Do
'If mytablex.EOF Then Exit Do
'If "" & mytablex.Fields("moneda") = "S" Then
'   xsoles = xsoles + Val("" & mytablex.Fields("total"))
'End If
'If "" & mytablex.Fields("moneda") = "D" Then
'   xdolares = xdolares + Val("" & mytablex.Fields("total"))
'End If
'soles = Format(xsoles, "0.00")
'dolares = Format(xdolares, "0.00")
'mytablex.MoveNext
'Loop
'Exit Sub
'cmd345_err:
'MsgBox "Aviso en sumar recibos " + error$, 48, "Aviso"
'Exit Sub
'End Sub

Sub sumar_recibos(mytablex As ADODB.Recordset)

    Dim xsoles   As Double

    Dim xdolares As Double

    On Error GoTo cmd345_err

    xsoles = 0
    xdolares = 0
    soles = "0.00"
    dolares = "0.00"

    xsolesanu = 0
    xdolaresanu = 0
    solesanu = "0.00"
    dolaresanu = "0.00"

    Do

        If mytablex.EOF Then Exit Do

        'Activos
        If "" & mytablex.Fields("moneda") = "S" And "" & mytablex.Fields("estado") = "2" Then
            xsoles = xsoles + Val("" & mytablex.Fields("total"))

        End If

        If "" & mytablex.Fields("moneda") = "D" And "" & mytablex.Fields("estado") = "2" Then
            xdolares = xdolares + Val("" & mytablex.Fields("total"))

        End If

        'Activos

        'anulados
        If "" & mytablex.Fields("moneda") = "S" And "" & mytablex.Fields("estado") = "1" Then
            xsolesanu = xsolesanu + Val("" & mytablex.Fields("total"))

        End If

        If "" & mytablex.Fields("moneda") = "D" And "" & mytablex.Fields("estado") = "1" Then
            xdolaresanu = xdolaresanu + Val("" & mytablex.Fields("total"))

        End If

        'anulados

        soles = Format(xsoles, "0.00")
        dolares = Format(xdolares, "0.00")

        solesanu = Format(xsolesanu, "0.00")
        dolaresanu = Format(xdolaresanu, "0.00")

        mytablex.MoveNext
    Loop
    Exit Sub
cmd345_err:
    MsgBox "Aviso en sumar recibos " + error$, 48, "Aviso"
    Exit Sub

End Sub

'' 30/11/2017 Correción  General del Sistema Parte I

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

    xtipoarchivo = xcuentaco
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
            'found = proceso_formatos(archivo_formato, mytablex, "/", "\", xtipoarchivo, "tmpcta", bxlocal, bxtipo, bxserie, bxnumero, ascopia, contando)
            found = proceso_formatos(archivo_formato, mytablex, "/", "\", xtipoarchivo, "tmpcta", bxlocal, bxtipo, bxserie, bxnumero, ascopia, contando)
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
    'found = proceso_formatos(archivo_formato, mytabley, "$", "?", "recibo", "recibo", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    found = proceso_formatos(archivo_formato, mytabley, "$", "?", "recibo", "recibo", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
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
        .Range(.Cells(4, 1), .Cells(4, 12)).Borders.LineStyle = xlContinuous
        .Range(.Cells(4, 1), .Cells(4, 12)).Font.bold = True
        
        .Range(.Cells(4, 1), .Cells(4, 12)).Interior.color = RGB(192, 192, 250)
        
        For I = 1 To Num_Campos Step 1
            .Cells(4, I) = Nombre_Campos(I)
        Next I

        'hasta aki pa colocar los titulos
        
        'a partir de aki ta claro que es pa darle el ancho a las celdas ;-)
        
        .columns("A").ColumnWidth = 5
        .columns("B").ColumnWidth = 5
        .columns("C").ColumnWidth = 5
        .columns("D").ColumnWidth = 5
        .columns("E").ColumnWidth = 12
        .columns("F").ColumnWidth = 10
        .columns("G").ColumnWidth = 15
        .columns("H").ColumnWidth = 20
        .columns("I").ColumnWidth = 3
        
        If Trim("" & rconsulta.Fields("acu")) = "W" Then
            .columns("J").ColumnWidth = 10
            .columns("K").ColumnWidth = 0

        End If
        
        If Trim("" & rconsulta.Fields("acu")) = "V" Then
            .columns("J").ColumnWidth = 0
            .columns("K").ColumnWidth = 10

        End If
           
        .columns("L").ColumnWidth = 30
        
    End With

End Function

Private Sub Reporte_Click()

    Dim v, h As Long

    Dim found       As Integer

    Dim I           As Integer

    Dim R           As Long

    Dim sdx         As Double
 
    Dim sdx1        As Double

    Dim sdx2        As Double

    Dim sdx3        As Double

    Dim sdx4        As Double
 
    Dim sdx1anu     As Double

    Dim sdx2anu     As Double

    Dim sdx3anu     As Double

    Dim sdx4anu     As Double
 
    Dim xingreso    As Double

    Dim xegreso     As Double
 
    Dim xingresoanu As Double

    Dim xegresoanu  As Double
    
    Dim Heading(13) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd1561212_err

    If MsgBox("Desea Generar reporte ", 1, "Aviso") <> 1 Then Exit Sub
    If rconsulta.RecordCount = 0 Then Exit Sub
    rconsulta.MoveFirst
   
    Heading(1) = "Estado"
    Heading(2) = "Local"
    Heading(3) = "Tipo"
    Heading(4) = "Serie"
    Heading(5) = "Numero"
    Heading(6) = "Fecha"
    Heading(7) = "Codigo"
    Heading(8) = "Nombre"
    Heading(9) = "M"
    Heading(10) = "Ingreso"
    Heading(11) = "Egreso"
    Heading(12) = "Observación"
   
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_Excelre(13, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
    v = 5
    h = 1
    sdx = 0
    sdx1 = 0
    sdx2 = 0
    sdx3 = 0
    sdx4 = 0

    sdxanu = 0
    sdx1anu = 0
    sdx2anu = 0
    sdx3anu = 0
    sdx4anu = 0

    If explreci.Caption = "INGRESO DINERO" Then
        objExcel.ActiveSheet.Cells(1, 6) = "   REPORTE DE INGRESOS DE DINERO"
    ElseIf explreci.Caption = "EGRESO DINERO" Then
        objExcel.ActiveSheet.Cells(1, 6) = "   REPORTE DE EGRESOS DE DINERO"

    End If
   
    objExcel.ActiveSheet.Cells(1, 6).Font.bold = True
    objExcel.ActiveSheet.Cells(1, 6).Font.Size = 14
    objExcel.ActiveSheet.Cells(1, 6).Font.color = RGB(0, 112, 184)
    
    objExcel.ActiveSheet.Cells(2, 1) = "FECHA INICIO  " + fechai
    objExcel.ActiveSheet.Cells(2, 5) = "FECHA FIN  " + fechaf
    
    'v = v + 1
    ' objExcel.ActiveSheet.Cells(v, h + 1) = "FechaI:" & fechai & " Fechaf:" & fechaf
    'v = v + 1

    Do

        If rconsulta.EOF Then Exit Do
            
        objExcel.ActiveSheet.Cells(v, h) = "'" & rconsulta.Fields("estado")
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & rconsulta.Fields("local")
        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & rconsulta.Fields("tipo")
        objExcel.ActiveSheet.Cells(v, h + 3) = "'" & rconsulta.Fields("serie")
        objExcel.ActiveSheet.Cells(v, h + 4) = "'" & rconsulta.Fields("numero")
        objExcel.ActiveSheet.Cells(v, h + 5) = "'" & rconsulta.Fields("fecha")
        objExcel.ActiveSheet.Cells(v, h + 6) = "'" & rconsulta.Fields("codigo")
        objExcel.ActiveSheet.Cells(v, h + 7) = "'" & rconsulta.Fields("nombre")
        objExcel.ActiveSheet.Cells(v, h + 8) = "'" & rconsulta.Fields("moneda")
            
        xingreso = 0
        xegreso = 0
        xingresoanu = 0
        xegresoanu = 0
            
        If Trim("" & rconsulta.Fields("acu")) = "W" Then
            xingreso = Val("" & rconsulta.Fields("total"))

        End If
            
        If Trim("" & rconsulta.Fields("acu")) = "V" Then
            xegreso = Val("" & rconsulta.Fields("total"))

        End If
            
        objExcel.ActiveSheet.Cells(v, h + 9) = xingreso
        objExcel.ActiveSheet.Cells(v, h + 10) = xegreso
        objExcel.ActiveSheet.Cells(v, h + 11) = "'" & rconsulta.Fields("observa")
            
        v = v + 1
            
        If Trim("" & rconsulta.Fields("moneda")) = "S" Then
                
            If Trim("" & rconsulta.Fields("acu")) = "W" And "" & rconsulta.Fields("estado") = "2" Then
                sdx1 = sdx1 + Val("" & rconsulta.Fields("total"))

            End If

            If Trim("" & rconsulta.Fields("acu")) = "W" And "" & rconsulta.Fields("estado") = "1" Then
                sdx1anu = sdx1anu + Val("" & rconsulta.Fields("total"))

            End If
   
            If Trim("" & rconsulta.Fields("acu")) = "V" And "" & rconsulta.Fields("estado") = "2" Then
                sdx2 = sdx2 + Val("" & rconsulta.Fields("total"))

            End If

            If Trim("" & rconsulta.Fields("acu")) = "V" And "" & rconsulta.Fields("estado") = "1" Then
                sdx2anu = sdx2anu + Val("" & rconsulta.Fields("total"))

            End If
            
        End If
            
        If Trim("" & rconsulta.Fields("moneda")) = "D" Then
                
            If Trim("" & rconsulta.Fields("acu")) = "W" And "" & rconsulta.Fields("estado") = "2" Then
                sdx3 = sdx3 + Val("" & rconsulta.Fields("total"))

            End If

            If Trim("" & rconsulta.Fields("acu")) = "W" And "" & rconsulta.Fields("estado") = "1" Then
                sdx3anu = sdx3anu + Val("" & rconsulta.Fields("total"))

            End If
                
            If Trim("" & rconsulta.Fields("acu")) = "V" And "" & rconsulta.Fields("estado") = "2" Then
                sdx4 = sdx4 + Val("" & rconsulta.Fields("total"))

            End If

            If Trim("" & rconsulta.Fields("acu")) = "V" And "" & rconsulta.Fields("estado") = "1" Then
                sdx4anu = sdx4anu + Val("" & rconsulta.Fields("total"))

            End If
            
        End If
            
        rconsulta.MoveNext
    Loop
     
    v = v + 1
            
    objExcel.ActiveSheet.Cells(v, h + 7) = "Total"
            
    objExcel.ActiveSheet.Cells(v, h + 8) = "S"
    objExcel.ActiveSheet.Cells(v + 1, h + 8) = "D"
            
    objExcel.ActiveSheet.Cells(v + 3, h + 7) = "Anulados"
    objExcel.ActiveSheet.Cells(v + 3, h + 7).Font.bold = True
            
    objExcel.ActiveSheet.Cells(v + 3, h + 8) = "S"
    objExcel.ActiveSheet.Cells(v + 4, h + 8) = "D"
            
    objExcel.ActiveSheet.Cells(v, h + 9) = sdx1 'total soles ingresos
    objExcel.ActiveSheet.Cells(v, h + 10) = sdx2 ' total soles egresos
            
    objExcel.ActiveSheet.Cells(v + 3, h + 9) = sdx1anu 'anulados soles ingresos
    objExcel.ActiveSheet.Cells(v + 3, h + 10) = sdx2anu ' anulados soles egresos
            
    Dim k As Integer

    For k = 8 To 11
        objExcel.ActiveSheet.Cells(v, k).Font.bold = True
        objExcel.ActiveSheet.Cells(v, k).Interior.color = RGB(248, 243, 53)
    Next

    ''
    'objExcel.ActiveSheet.Cells(v, h + 11) = sdx1 - sdx2
    objExcel.ActiveSheet.Cells(v, h + 11) = ""
    ''
            
    v = v + 1
            
    objExcel.ActiveSheet.Cells(v, h + 9) = sdx3 'total dolares ingresos
    objExcel.ActiveSheet.Cells(v, h + 10) = sdx4 ' total dolares egresos
                    
    objExcel.ActiveSheet.Cells(v + 3, h + 9) = sdx3anu 'anulados dolares ingresos
    objExcel.ActiveSheet.Cells(v + 3, h + 10) = sdx4anu ' anulados dolares egresos
            
    objExcel.ActiveSheet.Cells(v, h + 10) = sdx3 - sdx4

    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    MsgBox "Proceso Terminado ", 48, "Aviso"
    Exit Sub
cmd1561212_err:
    MsgBox "Error en exportacion excell " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub tipoclie_Click()

    If tipoclie = "C" Or tipoclie = "V" Then
        xcuentaco = "cuentac"
        XCUENTACO1 = "cuentacd"

    End If

    If tipoclie = "P" Then
        xcuentaco = "cuentap"
        XCUENTACO1 = "cuentapd"

    End If

End Sub

Sub carga_subconcepto(buf As String)

    Dim mytablex As New ADODB.Recordset

    subconcepto.Clear
    subconcepto.AddItem "%"
    mytablex.Open "select * from subconcepto where concepto='" & buf & "'", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        subconcepto.AddItem "" & mytablex.Fields("subconcepto") & "|" & mytablex.Fields("DESCRIPCIO")
        mytablex.MoveNext
    Loop
    mytablex.Close
    subconcepto.ListIndex = 0

End Sub
