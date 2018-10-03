VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmActualizaEstados 
   Caption         =   "Actualiza Estados Sunat"
   ClientHeight    =   8505
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14580
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
   ScaleHeight     =   8505
   ScaleWidth      =   14580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   1395
      Left            =   0
      ScaleHeight     =   1335
      ScaleWidth      =   14520
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   14580
      Begin VB.TextBox respuesta 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   5280
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "FrmActualizaEstados.frx":0000
         Top             =   120
         Width           =   4815
      End
      Begin VB.ComboBox local1 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   120
         Width           =   1815
      End
      Begin VB.TextBox fechaf 
         Height          =   375
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox fechai 
         Height          =   375
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
      Begin ChamaleonButton.ChameleonBtn BtnConsulta 
         Height          =   585
         Left            =   3120
         TabIndex        =   12
         Top             =   120
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   1032
         BTYPE           =   4
         TX              =   "Actualiza Estados del Portal"
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
         MICON           =   "FrmActualizaEstados.frx":0006
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RESULTADO >>>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3120
         TabIndex        =   11
         Top             =   840
         Width           =   1860
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Fin"
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid DgvBajas 
      Height          =   6300
      Left            =   16200
      TabIndex        =   6
      Top             =   4920
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
      ColumnCount     =   35
      BeginProperty Column00 
         DataField       =   "Estado"
         Caption         =   "E"
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
      BeginProperty Column11 
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
      BeginProperty Column12 
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
      BeginProperty Column18 
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
      BeginProperty Column19 
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
      BeginProperty Column20 
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
      BeginProperty Column21 
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
      BeginProperty Column22 
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
      BeginProperty Column23 
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
      BeginProperty Column24 
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
      BeginProperty Column25 
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
      BeginProperty Column26 
         DataField       =   "Neto"
         Caption         =   "Neto"
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
      BeginProperty Column27 
         DataField       =   "Descuento"
         Caption         =   "Descuento"
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
      BeginProperty Column28 
         DataField       =   "Subtotal"
         Caption         =   "Subtotal"
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
      BeginProperty Column29 
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
      BeginProperty Column30 
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
      BeginProperty Column31 
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
      BeginProperty Column32 
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
      BeginProperty Column33 
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
      BeginProperty Column34 
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
            ColumnWidth     =   209.764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2160
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   225.071
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   390.047
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   195.024
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column16 
            Object.Visible         =   0   'False
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column17 
            Object.Visible         =   0   'False
            ColumnWidth     =   345.26
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   569.764
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   615.118
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   2009.764
         EndProperty
         BeginProperty Column22 
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column25 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column26 
         EndProperty
         BeginProperty Column27 
         EndProperty
         BeginProperty Column28 
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
         EndProperty
         BeginProperty Column34 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DgvDoc 
      Height          =   6300
      Left            =   16080
      TabIndex        =   7
      Top             =   4800
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
      ColumnCount     =   35
      BeginProperty Column00 
         DataField       =   "Estado"
         Caption         =   "E"
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
      BeginProperty Column11 
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
      BeginProperty Column12 
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
      BeginProperty Column18 
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
      BeginProperty Column19 
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
      BeginProperty Column20 
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
      BeginProperty Column21 
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
      BeginProperty Column22 
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
      BeginProperty Column23 
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
      BeginProperty Column24 
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
      BeginProperty Column25 
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
      BeginProperty Column26 
         DataField       =   "Neto"
         Caption         =   "Neto"
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
      BeginProperty Column27 
         DataField       =   "Descuento"
         Caption         =   "Descuento"
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
      BeginProperty Column28 
         DataField       =   "Subtotal"
         Caption         =   "Subtotal"
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
      BeginProperty Column29 
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
      BeginProperty Column30 
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
      BeginProperty Column31 
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
      BeginProperty Column32 
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
      BeginProperty Column33 
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
      BeginProperty Column34 
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
            ColumnWidth     =   209.764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2160
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   225.071
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   390.047
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   195.024
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column16 
            Object.Visible         =   0   'False
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column17 
            Object.Visible         =   0   'False
            ColumnWidth     =   345.26
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   569.764
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   615.118
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   2009.764
         EndProperty
         BeginProperty Column22 
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column25 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column26 
         EndProperty
         BeginProperty Column27 
         EndProperty
         BeginProperty Column28 
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
         EndProperty
         BeginProperty Column34 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DgvTodo 
      Height          =   6300
      Left            =   0
      TabIndex        =   0
      Top             =   1800
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
Attribute VB_Name = "FrmActualizaEstados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rxconsultaTotal As New ADODB.Recordset

Dim rxconsultaDoc   As New ADODB.Recordset

Dim rxconsultaBajas As New ADODB.Recordset

'22/04/2018
Dim sdxtotal        As Double

Dim sdxtotal1       As Double

Dim sdxtotal2       As Double

'22/04/2018
Dim fechaResumen    As String

Sub sql_reporteFacturasNotas()

    On Error GoTo cmd37_err

    Dim buf As String

    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub
    buf = "select * from factura where"
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If local1 <> "%" Then
        buf = buf & " and local like '" & extra_loquesea(local1) & "'"

    End If

    buf = buf & " and SUBSTRING(SERIE,1,1)='F' and estado_sunat='PENDIENTE'  AND   estado_sunat<>'BAJA'  "

    buf = buf & " and (acu='A' OR acu='B' OR acu='C' OR acu='D' or acu='E' or acu='F') "

    If rxconsultaDoc.State = 1 Then rxconsultaDoc.Close
    rxconsultaDoc.Open buf, cn, adOpenStatic, adLockOptimistic

    If rxconsultaDoc.EOF = True And rxconsultaDoc.BOF = True Then

    End If
   
    Set DgvDoc.DataSource = rxconsultaDoc
    Exit Sub
cmd37_err:
    MsgBox "Error en select " & error$, 48, "Aviso"
    Exit Sub

End Sub

Sub sql_reporteBajas()

    On Error GoTo cmd37_err

    Dim buf As String

    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub
    buf = "select * from factura where"
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If local1 <> "%" Then
        buf = buf & " and local like '" & extra_loquesea(local1) & "'"

    End If

    buf = buf & " and SUBSTRING(SERIE,1,1)='F' and  estado_sunat='PENDIENTE_BAJA' AND ESTADO='1'  "
    buf = buf & " and (acu='A' OR acu='B' OR acu='C' OR acu='D' or acu='E' or acu='F') "

    If rxconsultaBajas.State = 1 Then rxconsultaBajas.Close
    rxconsultaBajas.Open buf, cn, adOpenStatic, adLockOptimistic

    If rxconsultaBajas.EOF = True And rxconsultaBajas.BOF = True Then

    End If
   
    Set DgvBajas.DataSource = rxconsultaBajas
    Exit Sub
cmd37_err:
    MsgBox "Error en select " & error$, 48, "Aviso"
    Exit Sub

End Sub

Sub sql_reporteTotal()

    On Error GoTo cmd37_err

    Dim buf As String

    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub
    buf = "select * from factura where"
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If local1 <> "%" Then
        buf = buf & " and local like '" & extra_loquesea(local1) & "'"

    End If

    buf = buf & " and (acu='A' OR acu='B' OR acu='C' OR acu='D' or acu='E' or acu='F') "

    If rxconsultaTotal.State = 1 Then rxconsultaTotal.Close
    rxconsultaTotal.Open buf, cn, adOpenStatic, adLockOptimistic

    If rxconsultaTotal.EOF = True And rxconsultaTotal.BOF = True Then

    End If
   
    Set DgvTodo.DataSource = rxconsultaTotal

    Exit Sub
cmd37_err:
    MsgBox "Error en select " & error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub BtnConsulta_Click()

    Dim v As Integer

    'Genera Txt Resumen de Boletas y Notas
    Call GeneraTxtResumenDiario

    'REsumen de Facturas
    sql_reporteFacturasNotas

    If rxconsultaDoc.RecordCount > 0 Then
        Do

            If rxconsultaDoc.EOF Then Exit Do
            my_local = (rxconsultaDoc.Fields("local"))
            my_tipo = (rxconsultaDoc.Fields("tipo"))
            my_acu = (rxconsultaDoc.Fields("acu"))
            my_serie = (rxconsultaDoc.Fields("serie"))
            my_numero = (rxconsultaDoc.Fields("numero"))
            DatosRptaXDocumento_sunat ' Facturas y sus Notas
            rxconsultaDoc.MoveNext
        Loop

    End If

    sql_reporteBajas

    If rxconsultaBajas.RecordCount > 0 Then
        Do

            If rxconsultaBajas.EOF Then Exit Do
            my_local = (rxconsultaBajas.Fields("local"))
            my_tipo = (rxconsultaBajas.Fields("tipo"))
            my_acu = (rxconsultaBajas.Fields("acu"))
            my_serie = (rxconsultaBajas.Fields("serie"))
            my_numero = (rxconsultaBajas.Fields("numero"))
            DatosRptaXBajas ' Facturas y sus Notas
            rxconsultaBajas.MoveNext
        Loop

    End If

    ' Facturas y sus Notas

    'Resumen de Boletas
    DatosRptaResumen_sunat
    sql_reporteTotal

End Sub

Private Sub Form_Activate()
    fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    fechaf = Format(Now, "dd/mm/yyyy")
    carga_inicial

End Sub

Function DatosRptaXDocumento_sunat()

    Dim salida           As Boolean

    Dim my_ruc           As String

    Dim file             As String

    Dim encontro         As Boolean

    Dim input_file       As String

    Dim hastaCuanto      As Integer

    Dim my_estadosunat   As String

    Dim my_numerointerno As String

    Dim my_tipointerno   As String

    Call Datos_Empresa(my_struc_datos_empresa(), my_local, salida, 0)
    my_ruc = my_struc_datos_empresa(0).codigo1
            
    ' file = my_ruc & "_01_F851-00005201" & ".INPUT.TXT"
    If my_acu = "D" Then
        my_tipointerno = "01"
    ElseIf my_acu = "E" Then
        my_tipointerno = "07"
    ElseIf my_acu = "F" Then
        my_tipointerno = "08"

    End If
             
    my_numerointerno = my_numero
                 
    hastaCuanto = 8 - Len(my_numerointerno) '**en la tabla
    my_numerointerno = my_numerointerno
    Call E_llenar_zero(hastaCuanto, my_numerointerno, my_numerointerno)
                
    file = my_ruc & "_" & my_tipointerno & "_" & my_serie & "-" & my_numerointerno & ".INPUT.TXT"
             
    input_file = "D:\ce_Input\PROCESADO\R_" & Left(file, (Len(file) - 10)) & ".txt"
    salida = FileExists(input_file)

    If salida = False Then
        'MsgBox "Documento NO EXISTE", vbCritical
        Exit Function
    Else
        Call verifica_estado_electronicoXDocumento(input_file)

        If salida = False Then
            MsgBox "ERROR EN ARCHIVO", vbCritical
            Exit Function

        End If

    End If

End Function

Function DatosRptaXBajas()

    Dim salida           As Boolean

    Dim my_ruc           As String

    Dim file             As String

    Dim encontro         As Boolean

    Dim input_file       As String

    Dim hastaCuanto      As Integer

    Dim my_estadosunat   As String

    Dim my_numerointerno As String

    Dim my_tipointerno   As String

    Call Datos_Empresa(my_struc_datos_empresa(), my_local, salida, 0)
    my_ruc = my_struc_datos_empresa(0).codigo1
            
    ' file = my_ruc & "_01_F851-00005201" & ".INPUT.TXT"
    If my_acu = "D" Then
        my_tipointerno = "01"
    ElseIf my_acu = "E" Then
        my_tipointerno = "07"
    ElseIf my_acu = "F" Then
        my_tipointerno = "08"

    End If
             
    my_numerointerno = my_numero
                 
    hastaCuanto = 8 - Len(my_numerointerno) '**en la tabla
    my_numerointerno = my_numerointerno
    Call E_llenar_zero(hastaCuanto, my_numerointerno, my_numerointerno)
                
    file = my_ruc & "_RA_" & my_serie & "-" & my_numerointerno & ".INPUT.TXT"
             
    input_file = "D:\ce_Input\PROCESADO\R_" & Left(file, (Len(file) - 10)) & ".txt"
    salida = FileExists(input_file)

    If salida = False Then
        'MsgBox "Documento NO EXISTE", vbCritical
        Exit Function
    Else
        Call verifica_estado_electronicoXDocumento(input_file)

        If salida = False Then
            MsgBox "ERROR EN ARCHIVO", vbCritical
            Exit Function

        End If

    End If

End Function

Function DatosRptaResumen_sunat()

    Dim salida           As Boolean

    Dim my_ruc           As String

    Dim file             As String

    Dim encontro         As Boolean

    Dim input_file       As String

    Dim numerodias       As Integer

    Dim my_estadosunat   As String

    Dim my_numerointerno As String

    my_local = "" & extra_loquesea(local1)
           
    Call Datos_Empresa(my_struc_datos_empresa(), my_local, salida, 0)
    my_ruc = my_struc_datos_empresa(0).codigo1
            
    fechaResumen = Format(fechai, "dd/mm/yyyy")
    numerodias = DateDiff("d", fechai, fechaf)
     
    For I = 0 To numerodias
              
        If I = 0 Then
            fechaResumen = fechaResumen
        Else
            fechaResumen = DateAdd("D", 1, fechaResumen)

        End If
        
        fecha = Format(fechaResumen, "ddmmyyyy")
        file = my_ruc & "_RC_" & fecha & ".INPUT.TXT"
            
        input_file = "D:\ce_Input\PROCESADO\R_" & Left(file, (Len(file) - 10)) & ".txt"
                             
        salida = FileExists(input_file)
                   
        If salida = False Then
            MsgBox "Falta RESUMEN DIARIO. Fecha: " & fechaResumen, vbCritical
        Else
            salida = verifica_respuestaResumen_Sunat(input_file, fechaResumen)

        End If

    Next

End Function

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

End Sub

Function GeneraTxtResumenDiario()

    Dim origene      As String

    Dim destinoe     As String

    Dim fechaResumen As String

    Dim numerodias   As Integer

    Dim I            As Integer

    Dim salida       As Boolean

    Dim fecha        As String

    my_local = "" & extra_loquesea(local1)
    Call Datos_Empresa(my_struc_datos_empresa(), my_local, salida, 0)
    my_ruc = my_struc_datos_empresa(0).codigo1
    
    fechaResumen = Format(fechai, "dd/mm/yyyy")
    numerodias = DateDiff("d", fechai, fechaf)

    For I = 0 To numerodias

        If I = 0 Then
            fechaResumen = fechaResumen
        Else
            fechaResumen = DateAdd("D", 1, fechaResumen)

        End If
        
        fecha = Format(fechaResumen, "ddmmyyyy")
        FileName = "D:\ce_output\" & my_ruc & "_RC_" & fecha & ".INPUT.TXT"
        file = my_ruc & "_RC_" & fecha & ".INPUT.TXT"
        Filelibero1 = FreeFile
        Open FileName For Append As #Filelibero1
        Close #Filelibero1
        
        origene = FileName
        destinoe = "D:\ce_Input\" & file
        
        Dim fso As New Scripting.FileSystemObject

        fso.MoveFile origene, destinoe

    Next

End Function

