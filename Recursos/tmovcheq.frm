VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tmovcheq 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimientos"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   15450
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   15450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   0
      TabIndex        =   81
      Top             =   0
      Visible         =   0   'False
      Width           =   13935
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
         MaxLength       =   10
         TabIndex        =   84
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
         TabIndex        =   83
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
         Left            =   10800
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dbgrid1 
         Height          =   6735
         Left            =   240
         TabIndex        =   85
         Top             =   1200
         Width           =   12135
         _ExtentX        =   21405
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
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Datos Movimientos"
      Height          =   8175
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Visible         =   0   'False
      Width           =   14775
      Begin VB.TextBox xbanco 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   62
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox XCUENTA 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   61
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox transaccio 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   60
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox neto 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   59
         Top             =   5160
         Width           =   1575
      End
      Begin VB.TextBox descuento 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   58
         Top             =   4800
         Width           =   1575
      End
      Begin VB.TextBox nnombre 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1680
         MaxLength       =   60
         TabIndex        =   57
         Top             =   3000
         Width           =   4575
      End
      Begin VB.TextBox codigo 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   56
         Top             =   2640
         Width           =   1575
      End
      Begin VB.ComboBox tipoclie 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   2280
         Width           =   1215
      End
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
         Height          =   1575
         Left            =   8880
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tmovcheq.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Grabar registro"
         Top             =   1200
         Width           =   2055
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
         Height          =   1575
         Left            =   8880
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tmovcheq.frx":1212
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Borrar registro"
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox importe 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   52
         Top             =   4440
         Width           =   1575
      End
      Begin VB.ComboBox concilia 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox fechae 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   50
         Top             =   3720
         Width           =   1575
      End
      Begin VB.TextBox fechan 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   49
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox numero 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   48
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox tipo 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   47
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox comenta 
         BackColor       =   &H00C0FFFF&
         Height          =   1695
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   46
         Top             =   5880
         Width           =   6975
      End
      Begin VB.TextBox abono 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   45
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Banco"
         Height          =   375
         Left            =   120
         TabIndex        =   80
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Correlativo"
         Height          =   375
         Left            =   120
         TabIndex        =   79
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NetoBanco"
         Height          =   375
         Left            =   120
         TabIndex        =   78
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DsctosBancarios"
         Height          =   375
         Left            =   120
         TabIndex        =   77
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         Height          =   375
         Left            =   120
         TabIndex        =   76
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   75
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoClie"
         Height          =   375
         Left            =   120
         TabIndex        =   74
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label acu 
         BackColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   6360
         TabIndex        =   73
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label descripcio 
         BackColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   3360
         TabIndex        =   72
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Importe"
         Height          =   375
         Left            =   120
         TabIndex        =   71
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Concilia"
         Height          =   375
         Left            =   120
         TabIndex        =   70
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaEfectiva"
         Height          =   375
         Left            =   120
         TabIndex        =   69
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaNominal"
         Height          =   375
         Left            =   120
         TabIndex        =   68
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero"
         Height          =   375
         Left            =   120
         TabIndex        =   67
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         Height          =   375
         Left            =   120
         TabIndex        =   66
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comentarios"
         Height          =   375
         Left            =   120
         TabIndex        =   65
         Top             =   5880
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuenta"
         Height          =   375
         Left            =   120
         TabIndex        =   64
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Abonos"
         Height          =   375
         Left            =   120
         TabIndex        =   63
         Top             =   5520
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid dbgrid2 
      Height          =   6495
      Left            =   120
      TabIndex        =   43
      Top             =   1320
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   11456
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   22
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   27
      BeginProperty Column00 
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
      BeginProperty Column01 
         DataField       =   "Banco"
         Caption         =   "Banco"
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
         DataField       =   "Cuenta"
         Caption         =   "Cuenta"
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
         DataField       =   "fechan"
         Caption         =   "Fechan"
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
         DataField       =   "Fechae"
         Caption         =   "Fechae"
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
         DataField       =   "tipoclie"
         Caption         =   "Tipoclie"
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
      BeginProperty Column07 
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
      BeginProperty Column08 
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
      BeginProperty Column09 
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
      BeginProperty Column10 
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
         DataField       =   "Descuento"
         Caption         =   "Descuento"
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
         DataField       =   "Neto"
         Caption         =   "Neto"
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
         DataField       =   "Abono"
         Caption         =   "Abono"
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
         DataField       =   "Saldo"
         Caption         =   "Saldo"
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
         DataField       =   "Moneda"
         Caption         =   "Moneda"
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
         DataField       =   "Conciliado"
         Caption         =   "Conciliado"
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
         DataField       =   "concepto"
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
      BeginProperty Column19 
         DataField       =   "Paridad"
         Caption         =   "Paridad"
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
         DataField       =   "Comenta"
         Caption         =   "Comenta"
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
         DataField       =   "Cajero"
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
      BeginProperty Column22 
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
      BeginProperty Column23 
         DataField       =   "Turno"
         Caption         =   "Turno"
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
         DataField       =   "xtipo"
         Caption         =   "xtipo"
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
         DataField       =   "xserie"
         Caption         =   "xserie"
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
         DataField       =   "xnumero"
         Caption         =   "xnumero"
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
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   629.858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2055.118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   150.236
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   3075.024
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   390.047
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   225.071
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column20 
         EndProperty
         BeginProperty Column21 
         EndProperty
         BeginProperty Column22 
         EndProperty
         BeginProperty Column23 
         EndProperty
         BeginProperty Column24 
         EndProperty
         BeginProperty Column25 
         EndProperty
         BeginProperty Column26 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox ordenado 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9600
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   720
      Width           =   1455
   End
   Begin VB.ComboBox xtipo 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   720
      Width           =   4215
   End
   Begin VB.CommandButton Command6 
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
      Left            =   13440
      MaskColor       =   &H00FFFFFF&
      Picture         =   "tmovcheq.frx":2424
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.TextBox ynombre 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   11880
      MaxLength       =   10
      TabIndex        =   17
      Text            =   "%"
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox ycodigo 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   11880
      MaxLength       =   11
      TabIndex        =   15
      Text            =   "%"
      Top             =   0
      Width           =   1455
   End
   Begin VB.TextBox fechaf 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   9600
      MaxLength       =   10
      TabIndex        =   13
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox fechai 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   9600
      MaxLength       =   10
      TabIndex        =   11
      Top             =   0
      Width           =   1455
   End
   Begin VB.ComboBox banco 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   720
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tmovcheq.frx":2BD2
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Consulta"
      Top             =   0
      Width           =   735
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
      Left            =   1440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tmovcheq.frx":3DE4
      Style           =   1  'Graphical
      TabIndex        =   7
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
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tmovcheq.frx":4FF6
      Style           =   1  'Graphical
      TabIndex        =   6
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
      Picture         =   "tmovcheq.frx":6208
      Style           =   1  'Graphical
      TabIndex        =   5
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
      Picture         =   "tmovcheq.frx":741A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Nuevo registro"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox cuenta 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   4560
      MaxLength       =   20
      TabIndex        =   1
      Text            =   "%"
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label25 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Orden"
      Height          =   375
      Left            =   8880
      TabIndex        =   42
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label41 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SALDO"
      Height          =   375
      Left            =   5760
      TabIndex        =   40
      Top             =   8640
      Width           =   1335
   End
   Begin VB.Label Label40 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EGRESOS"
      Height          =   375
      Left            =   5760
      TabIndex        =   39
      Top             =   8280
      Width           =   1335
   End
   Begin VB.Label Label39 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "INGRESOS"
      Height          =   375
      Left            =   5760
      TabIndex        =   38
      Top             =   7920
      Width           =   1335
   End
   Begin VB.Label S5 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   13320
      TabIndex        =   37
      Top             =   8640
      Width           =   1575
   End
   Begin VB.Label S4 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11760
      TabIndex        =   36
      Top             =   8640
      Width           =   1575
   End
   Begin VB.Label S3 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10200
      TabIndex        =   35
      Top             =   8640
      Width           =   1575
   End
   Begin VB.Label S2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8640
      TabIndex        =   34
      Top             =   8640
      Width           =   1575
   End
   Begin VB.Label S1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7080
      TabIndex        =   33
      Top             =   8640
      Width           =   1575
   End
   Begin VB.Label E5 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   13320
      TabIndex        =   32
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Label E4 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11760
      TabIndex        =   31
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Label E3 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10200
      TabIndex        =   30
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Label E2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8640
      TabIndex        =   29
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Label E1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7080
      TabIndex        =   28
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Label I5 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   13320
      TabIndex        =   27
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label I4 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11760
      TabIndex        =   26
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label I3 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10200
      TabIndex        =   25
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label I2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8640
      TabIndex        =   24
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label I1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7080
      TabIndex        =   23
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFC0&
      Caption         =   "T=(X Cargos Y Descargos)"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo"
      Height          =   375
      Left            =   3840
      TabIndex        =   21
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label24 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre"
      Height          =   375
      Left            =   11160
      TabIndex        =   18
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo"
      Height          =   375
      Left            =   11160
      TabIndex        =   16
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaF"
      Height          =   375
      Left            =   8880
      TabIndex        =   14
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaI"
      Height          =   375
      Left            =   8880
      TabIndex        =   12
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Si son depositos realizados deben cruzar con boletas o facturas"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   7920
      Width           =   5535
   End
   Begin VB.Label moneda 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   14520
      TabIndex        =   9
      Top             =   840
      Width           =   120
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cuenta"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Banco"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin VB.Menu dlkio232 
      Caption         =   "&Nuevo"
   End
   Begin VB.Menu dmos8 
      Caption         =   "&Modifica"
   End
   Begin VB.Menu dk323 
      Caption         =   "&Borrar"
   End
   Begin VB.Menu dk23231 
      Caption         =   "&Zoom"
   End
   Begin VB.Menu dimproe 
      Caption         =   "&Imprime"
   End
   Begin VB.Menu ldo23 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tmovcheq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txcheque As New ADODB.Recordset

Private Sub abono_KeyPress(KeyAscii As Integer)

    Dim sdx As Double

    If KeyAscii <> 13 Then Exit Sub
    sdx = Val(importe) - Val(descuento)
    neto = Format(sdx, "0.00")

    comenta.SetFocus

End Sub

Private Sub abono_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        descuento.SetFocus
        Exit Sub

    End If

End Sub

Private Sub banco_Click()

    'consulta_sql
End Sub

Private Sub banco_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    consulta_sql

End Sub

Private Sub banco_KeyUp(KeyCode As Integer, Shift As Integer)
    'If KeyCode = &H70 Then  'f1
    '   consulta_banco
    'End If

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        ldo23_Click
        Exit Sub

    End If

    Command1_Click

End Sub

Private Sub cmdAddEntry_Click()

    Dim found As Integer

    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    'If banco = "*" Then
    '   MsgBox "Seleccione un banco", 48, "Aviso"
    '   Exit Sub
    'End If
    consulta_sql
    tipo.Enabled = True
    Numero.Enabled = True
    Frame2.Caption = "NUEVO"
    Frame2.Visible = True
    inicializa_todo
    fechan = Format(Now, "dd/mm/yyyy")
    fechae = Format(Now, "dd/mm/yyyy")

    xbanco.SetFocus

End Sub

Private Sub cmdDelete_Click()

    On Error GoTo cmd3_err

    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    If MsgBox("Esta Seguro,Borra el " & "" & txcheque.Fields("tipo") & " " & txcheque.Fields("numero"), 1, "Aviso") <> 1 Then Exit Sub
    txcheque.Delete
    sumar_detalle
    Exit Sub
cmd3_err:
    Exit Sub

End Sub

Private Sub cmdSave_Click()

End Sub

Private Sub cmdExit_Click()
    ldo23_Click

End Sub

Private Sub cmdHelp_Click()

End Sub

Private Sub cmdPrint_Click()

    Dim found As Integer

    Dim buf   As String

    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0

    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_documento
    cuerpo_programa_documento
    '------------------------------------
    Close #1
    cerrar_archivo
    'genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
    found = valida_wordpad(FileName)

End Sub

Sub cabecera_documento()

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
    buf = "Reporte de Movimientos de Bancos "
    found = formateaa(buf, 90, 2, 0)
        
    buf = String(150, "-")
    found = formateaa(buf, 150, 2, 0)
    '------aqui van los registros----------------------
    found = formateaa("Banco  :", 10, 0, 0)
    found = formateaa(extra_loquesea(banco), 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = busca_xbanco(0)
    found = formateaa(buf, 15, 2, 0)
        
    found = formateaa("Cuenta :", 10, 0, 0)
    found = formateaa(cuenta, 20, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = busca_xcuenta("" & extra_loquesea(banco), "" & cuenta)
    found = formateaa(buf, 15, 2, 0)
    
    found = formateaa("FechaN", 11, 0, 0)
    found = formateaa("FechaE", 11, 0, 0)
    found = formateaa("tipo", 7, 0, 0)
    found = formateaa("Numero", 16, 0, 0)
    found = formateaa("Concepto", 16, 0, 0)
    found = formateaa("M", 2, 0, 0)
    found = formateaa("Debito ", 11, 0, 1)
    found = formateaa("Credito ", 11, 0, 1)
    found = formateaa("Dsctos ", 11, 0, 1)
    found = formateaa("Comentarios", 40, 2, 0)
    
    '--------------------------------------------------
    buf = String(150, "-")
    found = formateaa(buf, 150, 2, 0)

End Sub

Sub cuerpo_programa_documento()

    Dim debito  As Double

    Dim credito As Double

    Dim buf     As String

    Dim found   As Integer

    Dim sdx     As Double

    Dim sdx1    As Double

    On Error GoTo cmd788_err

    sdx = 0
    sdx1 = 0
    debito = 0
    credito = 0
    ir_inicio
    Do

        If txcheque.EOF Then Exit Do
        '-----------------------------------------
        buf = "" & txcheque.Fields("fechaN")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txcheque.Fields("fechae")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txcheque.Fields("tipo")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txcheque.Fields("numero")
        found = formateaa(buf, 15, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txcheque.Fields("concepto")
        found = formateaa(buf, 15, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txcheque.Fields("moneda")
        found = formateaa(buf, 1, 0, 0)
        found = formateaa("", 1, 0, 0)
        debito = 0
        credito = 0

        If "" & txcheque.Fields("acu") = "X" Then
            debito = Val("" & txcheque.Fields("neto"))

        End If

        If "" & txcheque.Fields("acu") = "Y" Then
            credito = Val("" & txcheque.Fields("neto"))

        End If

        buf = "" & debito
        buf = Format(Val(buf), "0.00")

        If Val(buf) <= 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & credito
        buf = Format(Val(buf), "0.00")

        If Val(buf) <= 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & txcheque.Fields("descuento")
        buf = Format(Val(buf), "0.00")

        If Val(buf) <= 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & txcheque.Fields("comenta")
        found = formateaa(buf, 39, 0, 0)
        found = formateaa("", 1, 2, 0)
        nlineas
        sdx = sdx + debito
        sdx1 = sdx1 + credito
        '-----------------------------------------
        txcheque.MoveNext
    Loop
    found = formateaa("", 63, 0, 0)
    buf = Format(sdx, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(sdx1, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)

    buf = "Neto Bruto Bancos :" & Format(sdx - sdx1, "0.00")
    found = formateaa(buf, 25, 0, 1)
    found = formateaa("", 1, 2, 0)
    Exit Sub
cmd788_err:
    Exit Sub

End Sub

Sub nlineas()
    contlin = contlin + 1

    If contlin > 45 Then
        cabecera_documento

    End If

End Sub

Private Sub cmdSort_Click()

    Dim found As Integer

    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    'If banco = "*" Then
    '   MsgBox "Seleccione un Banco", 48, "Aviso"
    '   banco.SetFocus
    '   Exit Sub
    'End If
    found = pone_registro()

    If found = 0 Then
        MsgBox "Seleccione Un registro", 48, "Aviso"
        Exit Sub

    End If

    Command3.Enabled = True
    Frame2.Caption = "MODIFICA"
    Frame2.Visible = True
    tipo.Enabled = True
    Numero.Enabled = True
    transaccio.Enabled = False
    xbanco.SetFocus

End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(codigo) > 0 Then
        found = busca_codigo()

    End If

    nnombre.SetFocus

End Sub

Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_cliente

    End If

    If KeyCode = &H26 Then
        If Numero.Enabled = False Then Exit Sub
        tipoclie.SetFocus
        Exit Sub

    End If

    If KeyCode = &H76 Then  'f7
        If tipoclie = "C" Then

            'tnclie.show 1
        End If

        If tipoclie = "P" Then

            'tnprov.show 1
        End If

        If tipoclie = "V" Then
            tpersona.Show 1

        End If
   
    End If

End Sub

Private Sub comenta_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    Command3_Click

End Sub

Private Sub comenta_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        abono.SetFocus
        Exit Sub

    End If

End Sub

Private Sub Command1_Click()

    Dim bufxx     As String

    Dim buf       As String

    Dim txcheque1 As New ADODB.Recordset

    If tipoclie = "C" Then
        bufxx = "clientes"

    End If

    If tipoclie = "V" Then
        bufxx = "VENDEDOR"

    End If

    If tipoclie = "P" Then
        bufxx = "PROVEEDO"

    End If

    If opcion1 = "13" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from  " & bufxx
        Else
            buf = "select Nombre,Codigo from " & bufxx & " where " & Combo1 & " like '" & buffer & "%'"

        End If
   
    End If

    If opcion1 = "12" Then
        If Len(buffer) = 0 Then
            buf = "select Descripcio,Banco from banco "
        Else
            buf = "select Descripcio,Banco from banco where " & Combo1 & " like '" & buffer & "%'"

        End If
   
    End If

    If opcion1 = "1" Then
        If Len(buffer) = 0 Then
            buf = "select Descripcio,Cuenta,Banco from tabbanco where banco='" & extra_loquesea(banco) & "'"
        Else
            buf = "select Descripcio,Cuenta,Banco from tabbanco where banco='" & extra_loquesea(banco) & "' and " & Combo1 & " like '" & buffer & "%'"

        End If
  
    End If

    If opcion1 = "20" Then
        If Len(buffer) = 0 Then
            buf = "select Descripcio,Cuenta,Banco from tabbanco where banco='" & xbanco & "'"
        Else
            buf = "select Descripcio,Cuenta,Banco from tabbanco where banco='" & xbanco & "' and " & Combo1 & " like '" & buffer & "%'"

        End If
   
    End If

    If opcion1 = "2" Then
        If Len(buffer) = 0 Then
            buf = "select Descripcio,Tipo from tipo where tipodoc='X' or tipodoc='Y'"
        Else
            buf = "select Descripcio,Tipo from tipo where (tipodoc='X' or tipodoc='Y') and " & Combo1 & " like '" & buffer & "%'"

        End If
   
    End If

    If txcheque1.State = 1 Then txcheque1.Close
    txcheque1.Open buf, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = txcheque1

    dbGrid1.columns(0).Width = 4000
    dbGrid1.columns(1).Width = 2000

    If txcheque1.RecordCount > 0 Then
        dbGrid1.SetFocus

    End If

End Sub

Private Sub Command2_Click()
    ldo23_Click

End Sub

Private Sub Command3_Click()

    Dim found As Integer

    Dim sdx   As Double

    sdx = Val(importe) - Val(descuento)
    neto = Format(sdx, "0.00")

    If Len(xbanco) = 0 Then
        xbanco.SetFocus
        Exit Sub

    End If

    If Len(XCUENTA) = 0 Then
        XCUENTA.SetFocus
        Exit Sub

    End If

    found = valida_cuentas()

    If found = 0 Then
        MsgBox "No existe banco o cuenta", 48, "Aviso"
        XCUENTA.SetFocus

    End If

    If Len(tipo) = 0 Then
        tipo.SetFocus
        Exit Sub

    End If

    If Len(Numero) = 0 Then
        Numero.SetFocus
        Exit Sub

    End If

    If Not IsDate(fechan) Then
        fechan.SetFocus
        Exit Sub

    End If

    If Not IsDate(fechae) Then
        fechae.SetFocus
        Exit Sub

    End If

    If Val(importe) = 0 Then
        importe.SetFocus
        Exit Sub

    End If

    found = busca_tipo()

    If found = 0 Then
        tipo.SetFocus
        Exit Sub

    End If

    If Val(importe) = 0 Then
        importe.SetFocus
        Exit Sub

    End If

    sdx = 0
    sdx = Val(importe) - Val(descuento)
    neto = Format(sdx, "0.00")

    If Frame2.Caption = "NUEVO" Then
        sdx = busca_general()
        sdx = sdx + 1
busca_n:
        transaccio = "" & sdx
        found = busca_movi(0)

        If found = 1 Then
            sdx = Val(transaccio) + 1
            GoTo busca_n
            Exit Sub

        End If

        txcheque.AddNew
        graba_registro
        txcheque.Update

        If Len(codigo) > 0 And Len(nnombre) > 0 Then
            graba_cliente

        End If

    End If

    If Frame2.Caption = "MODIFICA" Then
        'txcheque.Edit
        graba_registro
        txcheque.Update

    End If

    ldo23_Click
    sumar_detalle

    'Command2_Click
End Sub

Function pone_registro()

    On Error GoTo cmd2_err

    xbanco = "" & txcheque.Fields("banco")
    XCUENTA = "" & txcheque.Fields("cuenta")
    transaccio = "" & txcheque.Fields("transaccio")
    nnombre = "" & txcheque.Fields("nombre")
    comenta = "" & txcheque.Fields("comenta")
    descuento = "" & txcheque.Fields("descuento")
    codigo = "" & txcheque.Fields("codigo")
    tipo = "" & txcheque.Fields("tipo")
    Numero = "" & txcheque.Fields("numero")
    fechan = "" & txcheque.Fields("fechan")
    fechae = "" & txcheque.Fields("fechae")
    abono = "" & txcheque.Fields("abono")
    concilia.ListIndex = 0

    If "" & txcheque.Fields("conciliado") = "S" Then
        concilia.ListIndex = 1

    End If

    descripcio = "" & txcheque.Fields("concepto")
    importe = "" & txcheque.Fields("total")
    tipoclie.ListIndex = 0

    If "" & txcheque.Fields("tipoclie") = "P" Then '
        tipoclie.ListIndex = 1

    End If

    If "" & txcheque.Fields("tipoclie") = "V" Then '
        tipoclie.ListIndex = 2

    End If

    pone_registro = 1
    Exit Function
cmd2_err:
    MsgBox error$, 48, "Aviso"
    Exit Function

End Function

Sub graba_registro()

    Dim buf As String

    Dim sdx As Double

    buf = valida_bancos()

    If buf <> "S" And buf <> "D" Then
        buf = "S"

    End If

    sdx = Val(importe) - Val(descuento)
    neto = Format(sdx, "0.00")
    txcheque.Fields("transaccio") = transaccio
    txcheque.Fields("moneda") = buf
    txcheque.Fields("tipoclie") = tipoclie
    txcheque.Fields("codigo") = codigo
    txcheque.Fields("banco") = xbanco
    txcheque.Fields("cuenta") = XCUENTA
    txcheque.Fields("tipo") = tipo
    txcheque.Fields("numero") = Numero
    txcheque.Fields("fechan") = Format(fechan, "dd/mm/yyyy")
    txcheque.Fields("fechae") = Format(fechae, "dd/mm/yyyy")
    txcheque.Fields("nombre") = nnombre
    txcheque.Fields("conciliado") = concilia
    txcheque.Fields("concepto") = descripcio
    txcheque.Fields("acu") = acu
    txcheque.Fields("comenta") = comenta
    txcheque.Fields("total") = Val(importe)
    txcheque.Fields("descuento") = Val(descuento)
    txcheque.Fields("abono") = Val(abono)
    txcheque.Fields("recargo") = 0
    txcheque.Fields("neto") = Val(importe) - Val(descuento)
    txcheque.Fields("saldo") = Val(neto) - Val(abono)

End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command6_Click()

    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub

    consulta_sql

End Sub

Private Sub concepto_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    concilia.SetFocus

End Sub

Private Sub concepto_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        fechae.SetFocus
        Exit Sub

    End If

End Sub

Private Sub concilia_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    importe.SetFocus

End Sub

Private Sub concilia_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        fechae.SetFocus
        Exit Sub

    End If

End Sub

Private Sub cuenta_KeyPress(KeyAscii As Integer)

    Dim found As Integer

End Sub

Private Sub cuenta_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        If banco = "%" Then
            banco.SetFocus
            Exit Sub

        End If

        consulta_cuenta

    End If

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "20" Then
            XCUENTA = dbGrid1.columns(1)
            'banco = DBGrid1.Columns(2)
            Frame1.Visible = False
            Frame1.Enabled = False
            XCUENTA.SetFocus
            XCUENTA_KeyPress 13

        End If
   
        If opcion1 = "1" Then
            cuenta = dbGrid1.columns(1)
            'banco = DBGrid1.Columns(2)
            Frame1.Visible = False
            Frame1.Enabled = False
            cuenta.SetFocus
            cuenta_KeyPress 13

        End If

        If opcion1 = "13" Then
            codigo = dbGrid1.columns(1)
            Frame1.Visible = False
            Frame1.Enabled = False
            codigo.SetFocus
            codigo_KeyPress 13

        End If

        If opcion1 = "12" Then
            xbanco = dbGrid1.columns(1)
            Frame1.Visible = False
            Frame1.Enabled = False
            xbanco.SetFocus
            xbanco_KeyPress 13

        End If

        If opcion1 = "2" Then
            tipo = dbGrid1.columns(1)
            Frame1.Visible = False
            Frame1.Enabled = False
            tipo.SetFocus
            tipo_KeyPress 13

        End If
   
    End If

End Sub

Private Sub DBGrid2_DblClick()
    cmdSort_Click

End Sub

Private Sub dk882_Click()

End Sub

Private Sub descuento_KeyPress(KeyAscii As Integer)

    Dim sdx As Double

    If KeyAscii <> 13 Then Exit Sub
    sdx = Val(importe) - Val(descuento)
    neto = Format(sdx, "0.00")
    abono.SetFocus

End Sub

Private Sub descuento_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        importe.SetFocus
        Exit Sub

    End If

End Sub

Private Sub dimproe_Click()

    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub

    cmdPrint_Click

End Sub

Private Sub dk23231_Click()

    Dim found As Integer

    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub

    'If banco = "*" Then
    '   MsgBox "Seleccione un Banco", 48, "Aviso"
    '   banco.SetFocus
    '   Exit Sub
    'End If
    found = pone_registro()

    If found = 0 Then
        MsgBox "Seleccione Un registro", 48, "Aviso"
        Exit Sub

    End If

    Command3.Enabled = False
    Frame2.Caption = "VER"
    Frame2.Visible = True
    tipo.Enabled = True
    Numero.Enabled = True
    transaccio.Enabled = False
    fechan.SetFocus

End Sub

Private Sub dk323_Click()

    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub

    cmdDelete_Click

End Sub

Private Sub dlkio232_Click()

    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub

    cmdAddEntry_Click

End Sub

Private Sub dmos8_Click()

    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub

    cmdSort_Click

End Sub

Private Sub fechae_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(fechae) = 0 Then
        fechae = Format(Now, "dd/mm/yyyy")

    End If

    concilia.SetFocus

End Sub

Private Sub fechae_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        fechan.SetFocus
        Exit Sub

    End If

End Sub

Private Sub fechan_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(fechan) = 0 Then
        fechan = Format(Now, "dd/mm/yyyy")

    End If

    fechae.SetFocus

End Sub

Private Sub fechan_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        nnombre.SetFocus
        Exit Sub

    End If

End Sub

Private Sub Form_Activate()
    'consulta_sql

End Sub

Private Sub Form_Load()
    concilia.Clear
    concilia.AddItem "N"
    concilia.AddItem "S"
    concilia.ListIndex = 0
    tipoclie.Clear
    tipoclie.AddItem "C"
    tipoclie.AddItem "P"
    tipoclie.AddItem "V"
    tipoclie.ListIndex = 0
    fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    fechaf = Format(Now, "dd/mm/yyyy")
    carga_banco
    carga_tipo
    ordenado.Clear
    ordenado.AddItem "TRANSACCIO"
    ordenado.AddItem "codigo"
    ordenado.AddItem "nombre"
    ordenado.AddItem "fechan"
    ordenado.AddItem "tipo"
    ordenado.AddItem "numero"
    ordenado.AddItem "cajero"
    ordenado.AddItem "caja"
    ordenado.AddItem "turno"
    ordenado.AddItem "xtipo"
    ordenado.AddItem "xserie"
    ordenado.AddItem "xnumero"
    ordenado.ListIndex = 0
    consulta_sql

End Sub

Private Sub importe_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    descuento.SetFocus

End Sub

Private Sub importe_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        concilia.SetFocus
        Exit Sub

    End If

End Sub

Private Sub ldo23_Click()

    If Frame2.Visible = True Then
        If Frame1.Visible = False Then
            Frame2.Visible = False
            DBGrid2.SetFocus
            Exit Sub

        End If
   
        If opcion1 = "13" Then
            Frame1.Visible = False
            codigo.Enabled = True
            codigo.SetFocus
            Exit Sub

        End If

        If opcion1 = "2" Then
            Frame1.Visible = False
            tipo.SetFocus
            Exit Sub

        End If

        If opcion1 = "20" Then
            Frame1.Visible = False
            XCUENTA.SetFocus
            Exit Sub

        End If

        If opcion1 = "12" Then
            Frame1.Visible = False
            xbanco.SetFocus
            Exit Sub

        End If

        Exit Sub

    End If

    If Frame1.Visible = True Then
        If opcion1 = "12" Then
            Frame1.Visible = False
            'banco.SetFocus
            Exit Sub

        End If

        If opcion1 = "13" Then
            Frame1.Visible = False
            codigo.Enabled = True
            codigo.SetFocus
            Exit Sub

        End If

        If opcion1 = "1" Then
            Frame1.Visible = False
            cuenta.SetFocus
            Exit Sub

        End If

        If opcion1 = "2" Then
            Frame1.Visible = False
            tipo.SetFocus
            Exit Sub

        End If

    End If

    tmovcheq.Hide
    Unload tmovcheq

End Sub

Sub consulta_cuenta()
    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "Cuenta"
    Combo1.AddItem "Banco"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "1"
    Command1_Click

End Sub

Sub consulta_tipo()
    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "tipo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "2"
    Command1_Click

End Sub

Sub consulta_sql()

    Dim buf  As String

    Dim buf1 As String

    buf = "select * from chequemo where  "
    buf = buf & "  fechan>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fechan<='" & Format(fechaf, "YYYYMMDD") & "' "

    If banco <> "%" Then
        buf = buf & " and banco like '" & extra_loquesea(banco) & "'"

    End If

    If cuenta <> "%" Then
        buf = buf & " and cuenta='" & cuenta & "'"

    End If

    If xtipo <> "%" Then
        buf = buf & " and tipo like '" & extra_loquesea(xtipo) & "'"

    End If

    If ycodigo <> "%" Then
        buf = buf & " and codigo like '" & ycodigo & "'"

    End If

    If ynombre <> "%" Then
        buf = buf & " and nombre like '" & ynombre & "'"

    End If

    If ordenado = "TRANSACCIO" Then
        buf1 = " str(transaccio)"
    Else
        buf1 = ordenado & ",str(transaccio)"

    End If

    buf = buf & " order by " & buf1
    'MsgBox buf

    If txcheque.State = 1 Then txcheque.Close
    txcheque.Open buf, cn, adOpenStatic, adLockOptimistic
    Set DBGrid2.DataSource = txcheque
    sumar_detalle
    'Data2.Refresh

End Sub

Private Sub nnombre_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    fechan.SetFocus

End Sub

Private Sub numero_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(Numero) = 0 Then
        Numero.SetFocus
        Exit Sub

    End If

    tipoclie.SetFocus

End Sub

Private Sub numero_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        tipo.SetFocus
        Exit Sub

    End If

End Sub

Sub inicializa_todo()
    xbanco = ""
    XCUENTA = ""
    transaccio = ""
    descuento = ""
    codigo = ""
    nnombre = ""
    tipoclie.ListIndex = 0
    tipo = ""
    Numero = ""
    fechan = ""
    fechae = ""
    'concepto = ""
    importe = ""
    'abono = ""
    comenta = ""
    concilia.ListIndex = 0

End Sub

Private Sub tipo_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 13 Then
        found = busca_tipo()

        If found = 0 Then Exit Sub
        Numero.SetFocus
        Exit Sub

    End If

    KeyAscii = 0

End Sub

Private Sub tipo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        XCUENTA.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_tipo

    End If

End Sub

Function busca_tipo()

    Dim mytablex As New ADODB.Recordset

    acu = ""
    descripcio = ""

    mytablex.Open "select * from tipo where tipo='" & tipo & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        descripcio = "" & mytablex.Fields("descripcio")
        acu = "" & mytablex.Fields("tipodoc")
        busca_tipo = 1

    End If

    mytablex.Close

End Function

Function busca_movi(sw As Integer)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from chequemo where transaccio='" & transaccio & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_movi = 1

        If sw = 0 Then

        End If

        If sw = 1 Then

        End If

    End If

    mytablex.Close

End Function

Sub consulta_banco()
    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "Banco"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "12"
    Command1_Click

End Sub

Sub consulta_cliente()
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "13"
    Command1_Click

End Sub

Function busca_codigo()

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    If tipoclie = "C" Then
        buf = "clientes"

    End If

    If tipoclie = "P" Then
        buf = "proveedo"

    End If

    If tipoclie = "V" Then
        buf = "vendedor"

    End If

    mytablex.Open "select * from " & buf & "  where codigo='" & codigo & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        nnombre = "" & mytablex.Fields("nombre")
        busca_codigo = 1

    End If

    mytablex.Close

End Function

Sub graba_cliente()

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    If tipoclie = "C" Then
        buf = "clientes"

    End If

    If tipoclie = "P" Then
        buf = "proveedo"

    End If

    If tipoclie = "V" Then
        buf = "vendedor"

    End If

    mytablex.Open "select * from " & buf & "  where codigo='" & codigo & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        mytablex.Fields("codigo") = codigo
        mytablex.Fields("nombre") = nnombre
        mytablex.Update
    Else
        mytablex.AddNew
        mytablex.Fields("codigo") = codigo
        mytablex.Fields("nombre") = nnombre
        mytablex.Update

    End If

    mytablex.Close

End Sub

Private Sub tipoclie_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    codigo.SetFocus

End Sub

Private Sub tipoclie_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        Numero.SetFocus
        Exit Sub

    End If

End Sub

Sub sumar_detalle()

    On Error GoTo cmd35_err

    Dim sdx1  As Double

    Dim sdx2  As Double

    Dim sdx3  As Double

    Dim sdx4  As Double

    Dim sdx5  As Double

    Dim sumaa As Double

    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0
    suma6 = 0

    suma7 = 0
    suma8 = 0
    suma9 = 0
    sumaa = 0

    Do

        If txcheque.EOF Then Exit Do
        If "" & txcheque.Fields("acu") = "X" Then
            suma1 = suma1 + Val("" & txcheque.Fields("Total"))
            suma2 = suma2 + Val("" & txcheque.Fields("descuento"))
            suma3 = suma3 + Val("" & txcheque.Fields("neto"))
            suma4 = suma4 + Val("" & txcheque.Fields("abono"))
            suma5 = suma5 + Val("" & txcheque.Fields("saldo"))

        End If

        If "" & txcheque.Fields("acu") = "Y" Then
            suma6 = suma6 + Val("" & txcheque.Fields("Total"))
            suma7 = suma7 + Val("" & txcheque.Fields("descuento"))
            suma8 = suma8 + Val("" & txcheque.Fields("neto"))
            suma9 = suma9 + Val("" & txcheque.Fields("abono"))
            sumaa = sumaa + Val("" & txcheque.Fields("saldo"))

        End If

        txcheque.MoveNext
    Loop
    'Data2.Refresh
    I1 = Format(suma1, "0.00")
    I2 = Format(suma2, "0.00")
    I3 = Format(suma3, "0.00")
    I4 = Format(suma4, "0.00")
    I5 = Format(suma5, "0.00")

    E1 = Format(suma6, "0.00")
    E2 = Format(suma7, "0.00")
    E3 = Format(suma8, "0.00")
    E4 = Format(suma9, "0.00")
    E5 = Format(sumaa, "0.00")

    S1 = Format(suma1 - suma6, "0.00")
    S2 = Format(suma2 - suma7, "0.00")
    S3 = Format(suma3 - suma8, "0.00")
    S4 = Format(suma4 - suma9, "0.00")
    S5 = Format(suma5 - sumaa, "0.00")
    DBGrid2.Col = 0
    DBGrid2.Row = DBGrid2.VisibleRows - 1
    DBGrid2.SetFocus
    Exit Sub
cmd35_err:
    'MsgBox "Error " & Error$ & " " & fila, 24, "Aviso"
    Exit Sub

End Sub

Private Sub totalc_Click()
    sumar_detalle

End Sub

Sub ir_inicio()

    On Error GoTo cmd34_err

    txcheque.MoveFirst
    Exit Sub
cmd34_err:
    Exit Sub

End Sub

Function busca_xbanco(sw As Integer) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from tabbanco  where banco='" & extra_loquesea(banco) & "' and cuenta='" & cuenta & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If sw = 0 Then
            busca_xbanco = "" & mytablex.Fields("descripcio")

        End If

        If sw = 1 Then
            busca_xbanco = "" & mytablex.Fields("moneda")

        End If

    End If

    mytablex.Close

End Function

Function valida_bancos() As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from tabbanco  where banco='" & xbanco & "' and cuenta='" & XCUENTA & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        valida_bancos = "" & mytablex.Fields("moneda")

    End If

    mytablex.Close

End Function

Function busca_xcuenta(buf As String, buf1 As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from tabbanco  where banco='" & buf & "' and cuenta='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_xcuenta = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close

End Function

Sub carga_banco()

    Dim mytablex As New ADODB.Recordset

    banco.Clear
    banco.AddItem "%"

    mytablex.Open "select * from banco", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        banco.AddItem "" & mytablex.Fields("banco") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    mytablex.Close
    banco.ListIndex = 0

End Sub

Sub carga_tipo()

    Dim mytablex As New ADODB.Recordset

    xtipo.Clear
    xtipo.AddItem "%"
    mytablex.Open "select * from tipo", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("tipodoc") = "X" Or "" & mytablex.Fields("tipodoc") = "Y" Then
            xtipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

        End If

        mytablex.MoveNext
    Loop
    mytablex.Close
    xtipo.ListIndex = 0

End Sub

Private Sub xbanco_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    XCUENTA.SetFocus

End Sub

Private Sub xbanco_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_banco

    End If

End Sub

Private Sub XCUENTA_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    tipo.SetFocus

End Sub

Private Sub XCUENTA_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        xbanco.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        If Len(xbanco) = 0 Then
            xbanco.SetFocus
            Exit Sub

        End If

        xconsulta_cuenta

    End If

End Sub

Private Sub xtipo_keyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

    'xtipo_Click
End Sub

Function busca_general() As Double

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_general = Val("" & mytablex.Fields("banco"))

    End If

    mytablex.Close

End Function

Sub xconsulta_cuenta()
    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "Cuenta"
    Combo1.AddItem "Banco"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "20"
    Command1_Click

End Sub

Function valida_cuentas()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from tabbanco where banco='" & xbanco & "' and cuenta='" & XCUENTA & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        valida_cuentas = 1

    End If

    mytablex.Close

End Function
