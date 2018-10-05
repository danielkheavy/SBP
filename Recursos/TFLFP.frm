VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tflfp 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formas de Pago"
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
   Begin MSDataGridLib.DataGrid dbgrid3 
      Height          =   1575
      Left            =   120
      TabIndex        =   45
      Top             =   6960
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   2778
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      Caption         =   "Consulta"
      Height          =   3615
      Left            =   4080
      TabIndex        =   4
      Top             =   2160
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
         Picture         =   "TFLFP.frx":0000
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
         Picture         =   "TFLFP.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
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
      Begin VB.ComboBox ordenado 
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
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   720
         Width           =   1815
      End
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
         Height          =   735
         Left            =   13680
         MaskColor       =   &H00FFFFFF&
         Picture         =   "TFLFP.frx":0F5C
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1335
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
         Height          =   375
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   14
         Top             =   0
         Width           =   1815
      End
      Begin VB.TextBox fechaf 
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
         Picture         =   "TFLFP.frx":170A
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
         Picture         =   "TFLFP.frx":291C
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
         Picture         =   "TFLFP.frx":3B2E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprimir"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ordenado"
         Height          =   375
         Left            =   3120
         TabIndex        =   47
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fpago"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10560
         TabIndex        =   35
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ClieTipo"
         Height          =   375
         Left            =   6000
         TabIndex        =   31
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Turno"
         Height          =   495
         Left            =   8040
         TabIndex        =   29
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
         Height          =   375
         Left            =   10560
         TabIndex        =   27
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         Height          =   255
         Left            =   10560
         TabIndex        =   26
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Usuario"
         Height          =   375
         Left            =   8040
         TabIndex        =   22
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         Height          =   375
         Left            =   6000
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         Height          =   375
         Left            =   6000
         TabIndex        =   19
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   375
         Left            =   3120
         TabIndex        =   16
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   375
         Left            =   3120
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
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
         DataField       =   "Fpago"
         Caption         =   "Fpago"
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
      BeginProperty Column13 
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
      BeginProperty Column14 
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
      BeginProperty Column15 
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
      BeginProperty Column16 
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
      BeginProperty Column17 
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
      BeginProperty Column18 
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
      BeginProperty Column19 
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
      BeginProperty Column20 
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
            ColumnWidth     =   629.858
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   2264.882
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   390.047
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   195.024
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   269.858
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   2234.835
         EndProperty
      EndProperty
   End
   Begin VB.Label Label24 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo"
      Height          =   375
      Left            =   8400
      TabIndex        =   44
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label esoles 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11640
      TabIndex        =   43
      Top             =   7800
      Width           =   1695
   End
   Begin VB.Label edolares 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10080
      TabIndex        =   42
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label Label19 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Egresos"
      Height          =   375
      Left            =   8400
      TabIndex        =   41
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label isoles 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11640
      TabIndex        =   40
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label idolares 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10080
      TabIndex        =   39
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ingresos"
      Height          =   375
      Left            =   8400
      TabIndex        =   38
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Soles"
      Height          =   375
      Left            =   11640
      TabIndex        =   37
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dolares"
      Height          =   375
      Left            =   10080
      TabIndex        =   36
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label estaya 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   13680
      TabIndex        =   24
      Top             =   7920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label dolares 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10080
      TabIndex        =   10
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label soles 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11640
      TabIndex        =   9
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Label afecta 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   13680
      TabIndex        =   8
      Top             =   8160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label acu 
      BackColor       =   &H00E0E0E0&
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
   Begin VB.Menu dki9923 
      Caption         =   "&Consulta"
      Visible         =   0   'False
   End
   Begin VB.Menu dkj8933 
      Caption         =   "&Reporte"
   End
   Begin VB.Menu lfo3434 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tflfp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rconsulta As New ADODB.Recordset

Dim rfpago    As New ADODB.Recordset

Private Sub cmdCancelar_Click()
    lfo3434_Click

End Sub

Private Sub cmdDelete_Click()

    'dbo912_Click
End Sub

Private Sub cmdGrabar_Click()
    Command1_Click
    'sql_recibos
    'consulta_fpago
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
    'borra_tabla "drop tempnose"
    prepara_temporal
    sql_recibos
    consulta_fpago

End Sub

Private Sub concepto_Click()

End Sub

Private Sub dbo912_Click()

End Sub

Function borra_fpagov(xlocal As String, _
                      xtipo As String, _
                      xserie As String, _
                      xnumero As String)
    cn.Execute ("delete from fpagov where local='" & xlocal & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'")

End Function

Private Sub dfkl8823_Click()

End Sub

Private Sub dki9923_Click()
    Frame2.Visible = True
    fechai.SetFocus

End Sub

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
   
    Heading(1) = "Lo"
    Heading(2) = "Tipo"
    Heading(3) = "Serie"
    Heading(4) = "Numero"
    Heading(5) = "Fecha"
    Heading(6) = "Codigo"
    Heading(7) = "Nombre"
    Heading(8) = "M"
    Heading(9) = "Ingreso"
    Heading(10) = "Egreso"
   
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_Excelre(12, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
    v = 5
    h = 1
    sdx = 0
    sdx1 = 0
    sdx2 = 0
    sdx3 = 0
    sdx4 = 0

    objExcel.ActiveSheet.Cells(v, h + 1) = "Reporte de ingresos Egresos"
    v = v + 1
    objExcel.ActiveSheet.Cells(v, h + 1) = "FechaI:" & fechai & " Fechaf:" & fechaf
    v = v + 1

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
            
        '''09/08/2017 kenyo Arreglo Reporte Flujo de Dinero
            
        'If Trim("" & rconsulta.Fields("acu")) = "W" Or Trim("" & rconsulta.Fields("acu")) = "C" Then
        'xingreso = Val("" & rconsulta.Fields("total"))
        'End If
        '
        'If Trim("" & rconsulta.Fields("acu")) = "V" Then
        'xegreso = Val("" & rconsulta.Fields("total"))
        'End If
                                
        If Trim("" & rconsulta.Fields("acu")) = "W" Or Trim("" & rconsulta.Fields("acu")) = "A" Or Trim("" & rconsulta.Fields("acu")) = "B" Or Trim("" & rconsulta.Fields("acu")) = "C" Or Trim("" & rconsulta.Fields("acu")) = "D" Or Trim("" & rconsulta.Fields("acu")) = "E" Or Trim("" & rconsulta.Fields("acu")) = "G" Or Trim("" & rconsulta.Fields("acu")) = "N" Then
            xingreso = Val("" & rconsulta.Fields("total"))

        End If
                    
        If Trim("" & rconsulta.Fields("acu")) = "V" Or Trim("" & rconsulta.Fields("acu")) = "J" Or Trim("" & rconsulta.Fields("acu")) = "K" Or Trim("" & rconsulta.Fields("acu")) = "L" Or Trim("" & rconsulta.Fields("acu")) = "M" Or Trim("" & rconsulta.Fields("acu")) = "P" Then
            xegreso = Val("" & rconsulta.Fields("total"))

        End If
            
        '''09/08/2017 kenyo Arreglo Reporte Flujo de Dinero
            
        objExcel.ActiveSheet.Cells(v, h + 8) = xingreso
        objExcel.ActiveSheet.Cells(v, h + 9) = xegreso
            
        'objExcel.ActiveSheet.Cells(v, h + 10) = "'" & rconsulta.Fields("observa")
            
        v = v + 1
          
        '''09/08/2017 kenyo Arreglo Reporte Flujo de Dinero
           
        'If Trim("" & rconsulta.Fields("moneda")) = "S" Then
        'If Trim("" & rconsulta.Fields("acu")) = "W" Then
        'sdx1 = sdx1 + Val("" & rconsulta.Fields("total"))
        'End If
        'If Trim("" & rconsulta.Fields("acu")) = "V" Then
        'sdx2 = sdx2 + Val("" & rconsulta.Fields("total"))
        'End If
        'End If
        '
        'If Trim("" & rconsulta.Fields("moneda")) = "D" Then
        'If Trim("" & rconsulta.Fields("acu")) = "W" Then
        'sdx3 = sdx3 + Val("" & rconsulta.Fields("total"))
        'End If
        'If Trim("" & rconsulta.Fields("acu")) = "V" Then
        'sdx4 = sdx4 + Val("" & rconsulta.Fields("total"))
        'End If
        'End If

        If Trim("" & rconsulta.Fields("moneda")) = "S" Then
            If Trim("" & rconsulta.Fields("acu")) = "W" Or Trim("" & rconsulta.Fields("acu")) = "A" Or Trim("" & rconsulta.Fields("acu")) = "B" Or Trim("" & rconsulta.Fields("acu")) = "C" Or Trim("" & rconsulta.Fields("acu")) = "D" Or Trim("" & rconsulta.Fields("acu")) = "E" Or Trim("" & rconsulta.Fields("acu")) = "G" Or Trim("" & rconsulta.Fields("acu")) = "N" Then
                sdx1 = sdx1 + Val("" & rconsulta.Fields("total"))

            End If
                
            If Trim("" & rconsulta.Fields("acu")) = "V" Or Trim("" & rconsulta.Fields("acu")) = "J" Or Trim("" & rconsulta.Fields("acu")) = "K" Or Trim("" & rconsulta.Fields("acu")) = "L" Or Trim("" & rconsulta.Fields("acu")) = "M" Or Trim("" & rconsulta.Fields("acu")) = "P" Then
                sdx2 = sdx2 + Val("" & rconsulta.Fields("total"))

            End If
          
        End If
            
        If Trim("" & rconsulta.Fields("moneda")) = "D" Then
            
            If Trim("" & rconsulta.Fields("acu")) = "W" Or Trim("" & rconsulta.Fields("acu")) = "A" Or Trim("" & rconsulta.Fields("acu")) = "B" Or Trim("" & rconsulta.Fields("acu")) = "C" Or Trim("" & rconsulta.Fields("acu")) = "D" Or Trim("" & rconsulta.Fields("acu")) = "E" Or Trim("" & rconsulta.Fields("acu")) = "G" Or Trim("" & rconsulta.Fields("acu")) = "N" Then
                sdx3 = sdx3 + Val("" & rconsulta.Fields("total"))

            End If
            
            If Trim("" & rconsulta.Fields("acu")) = "V" Or Trim("" & rconsulta.Fields("acu")) = "J" Or Trim("" & rconsulta.Fields("acu")) = "K" Or Trim("" & rconsulta.Fields("acu")) = "L" Or Trim("" & rconsulta.Fields("acu")) = "M" Or Trim("" & rconsulta.Fields("acu")) = "P" Then
                sdx4 = sdx4 + Val("" & rconsulta.Fields("total"))

            End If
            
        End If
        
        '''09/08/2017 kenyo Arreglo Reporte Flujo de Dinero
            
        rconsulta.MoveNext
    Loop
     
    '''09/08/2017 kenyo Arreglo Reporte Flujo de Dinero
    'v = v + 1
    'objExcel.ActiveSheet.Cells(v, h + 6) = "Total"
    'objExcel.ActiveSheet.Cells(v, h + 8) = sdx1
    'objExcel.ActiveSheet.Cells(v, h + 9) = sdx2
    'objExcel.ActiveSheet.Cells(v, h + 10) = sdx1 - sdx2
    'v = v + 1
    'objExcel.ActiveSheet.Cells(v, h + 8) = sdx3
    'objExcel.ActiveSheet.Cells(v, h + 9) = sdx4
    'objExcel.ActiveSheet.Cells(v, h + 9) = sdx3 - sdx4
    'objExcel.ActiveSheet.Cells(v, h + 9) = sdx3 - sdx4
            
    v = v + 1
    objExcel.ActiveSheet.Cells(v, h + 6) = "Total"
    objExcel.ActiveSheet.Cells(v, h + 7) = "S"
    objExcel.ActiveSheet.Cells(v, h + 8) = sdx1
    objExcel.ActiveSheet.Cells(v, h + 9) = sdx2
    objExcel.ActiveSheet.Cells(v, h + 10) = sdx1 - sdx2
    v = v + 1
    objExcel.ActiveSheet.Cells(v, h + 7) = "D"
    objExcel.ActiveSheet.Cells(v, h + 8) = sdx3
    objExcel.ActiveSheet.Cells(v, h + 9) = sdx4
    objExcel.ActiveSheet.Cells(v, h + 10) = sdx3 - sdx4
            
    '''09/08/2017 kenyo Arreglo Reporte Flujo de Dinero

    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    MsgBox "Proceso Terminado ", 48, "Aviso"
    Exit Sub
cmd1561212_err:
    MsgBox "Error en exportacion excell " + error$, 48, "Aviso"
    Exit Sub

End Sub

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
    Command1_Click

    'sql_recibos
    'consulta_fpago
End Sub

Private Sub Form_Activate()

    If estaya = "" Then
        fechai = Format(Now, "dd/mm/yyyy")
        fechaf = Format(Now, "dd/mm/yyyy")
        carga_inicial
        'prepara_temporal
        Command1_Click
        'sql_recibos
        'consulta_fpago
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
        'If "" & mytablex.Fields("tipodoc") = "W" Or "" & mytablex.Fields("tipodoc") = "V" Then
        tipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")
        'End If
        mytablex.MoveNext
    Loop
    mytablex.Close
    tipo.ListIndex = 0

    tacu.Clear
    tacu.AddItem "%"
    mytablex.Open "select * from fpago ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        tacu.AddItem "" & mytablex.Fields("fpago") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    mytablex.Close
    tacu.ListIndex = 0

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

    'concepto.Clear
    'concepto.AddItem "%"
    'mytablex.Open "select * from concepto ", cn, adOpenStatic, adLockOptimistic
    'Do
    'If mytablex.EOF Then Exit Do
    'concepto.AddItem Trim("" & mytablex.Fields("concepto")) & "|" & mytablex.Fields("DESCRIPCIO")
    'mytablex.MoveNext
    'Loop
    'mytablex.Close
    'concepto.ListIndex = 0

    'subconcepto.Clear
    'subconcepto.AddItem "%"
    'subconcepto.ListIndex = 0

End Sub

Private Sub Form_Load()

    tipoclie.Clear
    tipoclie.AddItem "%"
    tipoclie.AddItem "C"
    tipoclie.AddItem "P"
    tipoclie.AddItem "V"
    tipoclie.ListIndex = 0

    ordenado.Clear
    ordenado.AddItem "%"
    ordenado.AddItem "Fpago"
    ordenado.AddItem "Tipo"
    ordenado.AddItem "Fecha"
    ordenado.AddItem "caja"
    ordenado.AddItem "Usuario"
    ordenado.AddItem "Turno"
    ordenado.AddItem "Codigo"
    ordenado.ListIndex = 0

End Sub

Private Sub lfo3434_Click()

    If Frame2.Visible = True Then
        Frame2.Visible = False
        Exit Sub

    End If

    tflfp.Hide
    Unload tflfp

End Sub

Sub sql_recibos()

    On Error GoTo cmd37_err

    Dim buf As String

    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub
    buf = "select * from fpagov where "
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

    If cajero <> "%" Then
        buf = buf & " and usuario='" & extra_loquesea(cajero) & "'"

    End If

    If tacu <> "%" Then
        buf = buf & " and fpago='" & extra_loquesea(tacu) & "'"

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

    If ordenado <> "%" Then
        'buf = buf & " and acu='" & acu & "'"
        buf = buf & " order by " & ordenado

    End If

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
        'VENTAS

        If "" & mytablex.Fields("acu") = "W" Or "" & mytablex.Fields("acu") = "A" Or "" & mytablex.Fields("acu") = "B" Or "" & mytablex.Fields("acu") = "C" Or "" & mytablex.Fields("acu") = "D" Or "" & mytablex.Fields("acu") = "G" Then
            acumula_fpago mytablex, 1

            If "" & mytablex.Fields("moneda") = "S" Then
                xisoles = xisoles + Val("" & mytablex.Fields("total"))

            End If

            If "" & mytablex.Fields("moneda") = "D" Then
                xidolares = xidolares + Val("" & mytablex.Fields("total"))

            End If

            'acumula_fpago mytablex
        End If

        'COMPRAS
        If "" & mytablex.Fields("acu") = "V" Or "" & mytablex.Fields("acu") = "J" Or "" & mytablex.Fields("acu") = "K" Or "" & mytablex.Fields("acu") = "L" Or "" & mytablex.Fields("acu") = "M" Or "" & mytablex.Fields("acu") = "P" Then
            acumula_fpago mytablex, -1

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

Public Function Formato_Excelre(Num_Campos As Integer, _
                                Nombre_Campos() As String) As Boolean

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
        .columns("A").ColumnWidth = 3
        .columns("B").ColumnWidth = 5
        .columns("C").ColumnWidth = 5
        .columns("D").ColumnWidth = 12
        .columns("E").ColumnWidth = 10
        .columns("F").ColumnWidth = 15
        .columns("G").ColumnWidth = 20
        .columns("H").ColumnWidth = 3
        .columns("i").ColumnWidth = 10
        .columns("j").ColumnWidth = 10

    End With

End Function

Sub prepara_temporal()

    On Error GoTo cmd7878_err

    borra_tabla "drop table tempnose"
    cn.Execute ("select * into tempnose from fpago ")
    Exit Sub
cmd7878_err:
    MsgBox "Aviso en prepara temporal " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub acumula_fpago(mytabley As ADODB.Recordset, signo As Double)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from tempnose where fpago='" & "" & mytabley.Fields("fpago") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        mytablex.Fields("total") = Val("" & mytablex.Fields("total")) + signo * Val("" & mytabley.Fields("total"))
        mytablex.Update
    Else
        mytablex.AddNew
        mytablex.Fields("total") = Val("" & mytablex.Fields("total")) + Val("" & mytabley.Fields("total"))
        mytablex.Update

    End If

End Sub

Sub consulta_fpago()

    If rfpago.State = 1 Then
        rfpago.Close

    End If

    rfpago.Open "select Fpago,Descripcio,Total from tempnose", cn, adOpenStatic, adLockOptimistic
    Set dbgrid3.DataSource = rfpago

End Sub

