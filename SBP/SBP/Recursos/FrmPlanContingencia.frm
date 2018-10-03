VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmPlanContingencia 
   Caption         =   "Form2"
   ClientHeight    =   7140
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14565
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   7140
   ScaleWidth      =   14565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   1275
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   14505
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   14565
      Begin VB.ComboBox local1 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox servicios 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   15360
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ComboBox estado_sunat 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   9600
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Numero 
         Height          =   375
         Left            =   9600
         MaxLength       =   11
         TabIndex        =   9
         Text            =   "%"
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox fechaf 
         Height          =   375
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox fechai 
         Height          =   375
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   7
         Top             =   0
         Width           =   1815
      End
      Begin VB.ComboBox tipo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   9600
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   0
         Width           =   2415
      End
      Begin VB.ComboBox tipoContingencia 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox ordenado 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   0
         Width           =   2055
      End
      Begin VB.CommandButton GeneraResumenBoleta 
         Caption         =   "gENERA TXT"
         Height          =   600
         Left            =   14640
         TabIndex        =   3
         Top             =   120
         Width           =   990
      End
      Begin VB.CommandButton Command1 
         Caption         =   "factu"
         Height          =   360
         Left            =   16080
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ir a Generar"
         Height          =   360
         Left            =   12480
         TabIndex        =   1
         Top             =   840
         Width           =   1575
      End
      Begin ChamaleonButton.ChameleonBtn BtnConsulta 
         Height          =   825
         Left            =   3480
         TabIndex        =   13
         Top             =   120
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   1455
         BTYPE           =   4
         TX              =   "Ver en Pantalla"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmPlanContingencia.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cta"
         Height          =   375
         Left            =   16800
         TabIndex        =   22
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado_Sunat"
         Height          =   375
         Left            =   8400
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
         Height          =   375
         Left            =   5400
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Número"
         Height          =   375
         Left            =   8400
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Fin"
         Height          =   375
         Left            =   0
         TabIndex        =   18
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   375
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoDocto"
         Height          =   375
         Left            =   8400
         TabIndex        =   16
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Contingencia"
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ordenado"
         Height          =   375
         Left            =   5400
         TabIndex        =   14
         Top             =   0
         Width           =   900
      End
   End
   Begin MSDataGridLib.DataGrid DgvTodo 
      Height          =   6300
      Left            =   0
      TabIndex        =   23
      Top             =   1320
      Width           =   20730
      _ExtentX        =   36565
      _ExtentY        =   11113
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   19
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
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   34
      BeginProperty Column00 
         DataField       =   "Estado"
         Caption         =   "E"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
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
      BeginProperty Column08 
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
      BeginProperty Column09 
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
      BeginProperty Column10 
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
      BeginProperty Column11 
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
      BeginProperty Column12 
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
      BeginProperty Column13 
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
      BeginProperty Column14 
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
      BeginProperty Column15 
         DataField       =   "Servicio"
         Caption         =   "S"
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
         DataField       =   "observa"
         Caption         =   "Observacion"
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
         DataField       =   "retipo1"
         Caption         =   "retipo1"
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
         DataField       =   "renumero3"
         Caption         =   "renumero3"
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
         DataField       =   "renumero1"
         Caption         =   "renumero1"
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
         DataField       =   "renumero2"
         Caption         =   "renumero2"
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
      BeginProperty Column21 
         DataField       =   "Subtotal"
         Caption         =   "Op.Gravadas"
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
         DataField       =   "Gravado"
         Caption         =   "Op.Exonerado"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column23 
         DataField       =   "Impuesto"
         Caption         =   "Impuesto"
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
      BeginProperty Column24 
         DataField       =   "Total"
         Caption         =   "Total"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column25 
         DataField       =   "tipo1"
         Caption         =   "Tipo1"
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
      BeginProperty Column26 
         DataField       =   "Serie1"
         Caption         =   "Serie1"
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
      BeginProperty Column27 
         DataField       =   "numero1"
         Caption         =   "Numero1"
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
      BeginProperty Column28 
         DataField       =   "estado_sunat"
         Caption         =   "Estado_Sunat"
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
      BeginProperty Column29 
         DataField       =   "Tipoimp"
         Caption         =   "Tipoimp"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column30 
         DataField       =   "Acu1"
         Caption         =   "Acu1"
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
      BeginProperty Column31 
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
      BeginProperty Column32 
         DataField       =   "Fechae"
         Caption         =   "FechaE"
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
      BeginProperty Column33 
         DataField       =   "Hora"
         Caption         =   "Hora"
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
            Object.Visible         =   0   'False
            ColumnWidth     =   195.024
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2099.906
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   -1  'True
            ColumnWidth     =   225.071
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   390.047
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
            ColumnWidth     =   195.024
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column15 
            Object.Visible         =   0   'False
            ColumnWidth     =   345.26
         EndProperty
         BeginProperty Column16 
            Object.Visible         =   0   'False
            ColumnWidth     =   2009.764
         EndProperty
         BeginProperty Column17 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column18 
            Object.Visible         =   0   'False
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column19 
            Object.Visible         =   0   'False
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column20 
            Object.Visible         =   0   'False
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column25 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column26 
            ColumnWidth     =   615.118
         EndProperty
         BeginProperty Column27 
            Object.Visible         =   -1  'True
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column28 
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column29 
         EndProperty
         BeginProperty Column30 
         EndProperty
         BeginProperty Column31 
         EndProperty
         BeginProperty Column32 
         EndProperty
         BeginProperty Column33 
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmPlanContingencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnConsulta_Click()

    Dim salida As Boolean
  
    my_local = "" & extra_loquesea(local1)
    Call Datos_Empresa(my_struc_datos_empresa(), my_local, salida, 0)
    my_ruc2 = my_struc_datos_empresa(0).codigo1
    
    'Call estrae_PlanContingencia(my_ruc, extra_loquesea(tipoContingencia))
End Sub

Private Sub Form_Activate()
    fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    fechaf = Format(Now, "dd/mm/yyyy")
    carga_inicial

End Sub

Sub carga_inicial()

    Dim mytablex As New ADODB.Recordset

    tipoContingencia.Clear
    tipoContingencia.AddItem "%"
    mytablex.Open "select * from  tipocontingencia ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        tipoContingencia.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("descripcion")
        mytablex.MoveNext
    Loop
    mytablex.Close
    tipoContingencia.ListIndex = 0

    local1.Clear
    local1.AddItem "%"
    mytablex.Open "select * from tlocal ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        local1.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    local1.ListIndex = 1

End Sub
