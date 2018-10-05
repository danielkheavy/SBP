VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tsiscont 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Voucher"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   15090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00808080&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   0
      TabIndex        =   93
      Top             =   0
      Visible         =   0   'False
      Width           =   13935
      Begin VB.CommandButton Command10 
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
         Left            =   5520
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox bufferx 
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
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   600
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid dbgrid2 
         Height          =   6735
         Left            =   240
         TabIndex        =   97
         Top             =   1080
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
      Caption         =   "Mas Rapido"
      Height          =   8535
      Left            =   0
      TabIndex        =   39
      Top             =   120
      Visible         =   0   'False
      Width           =   10095
      Begin VB.TextBox xparidad 
         Height          =   495
         Left            =   7680
         MaxLength       =   11
         TabIndex        =   88
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox xmoneda 
         Height          =   495
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   87
         Text            =   "S"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox xexportacion 
         Height          =   495
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   85
         Top             =   5040
         Width           =   1695
      End
      Begin VB.TextBox xigvexonerado 
         Height          =   495
         Left            =   3600
         MaxLength       =   10
         TabIndex        =   81
         Top             =   6120
         Width           =   1695
      End
      Begin VB.TextBox xigvexportacion 
         Height          =   495
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   80
         Top             =   6120
         Width           =   1695
      End
      Begin VB.TextBox xigvbasegravada 
         Height          =   495
         Left            =   120
         MaxLength       =   10
         TabIndex        =   79
         Top             =   6120
         Width           =   1695
      End
      Begin VB.TextBox xfecha 
         Height          =   495
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   58
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox xfechav 
         Height          =   495
         Left            =   4560
         MaxLength       =   10
         TabIndex        =   57
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox xtipo 
         Height          =   495
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   56
         Text            =   "01"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox xcodigo 
         Height          =   495
         Left            =   1440
         MaxLength       =   11
         TabIndex        =   55
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox xglosa 
         Height          =   495
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   54
         Top             =   1680
         Width           =   6255
      End
      Begin VB.TextBox xmonto1 
         Height          =   495
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   53
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox xmonto2 
         Height          =   495
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   52
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox xmonto3 
         Height          =   495
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   51
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox xcuenta1 
         Height          =   495
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   50
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox xcuenta2 
         Height          =   495
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   49
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox xcuenta3 
         Height          =   495
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   48
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox xbasegravada 
         Height          =   495
         Left            =   120
         MaxLength       =   10
         TabIndex        =   47
         Top             =   5040
         Width           =   1695
      End
      Begin VB.TextBox xexonerado 
         Height          =   495
         Left            =   3600
         MaxLength       =   10
         TabIndex        =   46
         Top             =   5040
         Width           =   1695
      End
      Begin VB.TextBox xinafecto 
         Height          =   495
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   45
         Top             =   5040
         Width           =   1695
      End
      Begin VB.TextBox xisc 
         Height          =   495
         Left            =   7080
         MaxLength       =   10
         TabIndex        =   44
         Top             =   5040
         Width           =   1695
      End
      Begin VB.TextBox xotros 
         Height          =   495
         Left            =   5400
         MaxLength       =   10
         TabIndex        =   43
         Top             =   6120
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Grabar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2760
         TabIndex        =   42
         Top             =   6840
         Width           =   2055
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5280
         TabIndex        =   41
         Top             =   6840
         Width           =   2055
      End
      Begin VB.TextBox xnumero 
         Height          =   495
         Left            =   4560
         MaxLength       =   11
         TabIndex        =   40
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label ncodigo 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3240
         TabIndex        =   91
         Top             =   1200
         Width           =   6495
      End
      Begin VB.Label Label35 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Cambio"
         Height          =   495
         Left            =   6360
         TabIndex        =   90
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label33 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Moneda"
         Height          =   495
         Left            =   6360
         TabIndex        =   89
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Exportacion"
         Height          =   495
         Left            =   1800
         TabIndex        =   86
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Igv"
         Height          =   495
         Left            =   3600
         TabIndex        =   84
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Igv"
         Height          =   495
         Left            =   1800
         TabIndex        =   83
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Igv"
         Height          =   495
         Left            =   120
         TabIndex        =   82
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Label Label15 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaDoc"
         Height          =   495
         Left            =   120
         TabIndex        =   78
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaVenc."
         Height          =   495
         Left            =   3240
         TabIndex        =   77
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label17 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         Height          =   495
         Left            =   120
         TabIndex        =   76
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label18 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         Height          =   495
         Left            =   120
         TabIndex        =   75
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label19 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Glosa"
         Height          =   495
         Left            =   120
         TabIndex        =   74
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label20 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monto 1:"
         Height          =   495
         Left            =   120
         TabIndex        =   73
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label21 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monto 2:"
         Height          =   495
         Left            =   120
         TabIndex        =   72
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label22 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monto 3:"
         Height          =   495
         Left            =   120
         TabIndex        =   71
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Montos"
         Height          =   495
         Left            =   1440
         TabIndex        =   70
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cta Contable"
         Height          =   495
         Left            =   3240
         TabIndex        =   69
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         Height          =   495
         Left            =   5040
         TabIndex        =   68
         Top             =   2280
         Width           =   4695
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "baseImpGravada"
         Height          =   495
         Left            =   120
         TabIndex        =   67
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Exonerado"
         Height          =   495
         Left            =   3600
         TabIndex        =   66
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inafecto"
         Height          =   495
         Left            =   5280
         TabIndex        =   65
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Isc"
         Height          =   495
         Left            =   7080
         TabIndex        =   64
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "OtrosTrib."
         Height          =   495
         Left            =   5400
         TabIndex        =   63
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Label xncuenta1 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   5040
         TabIndex        =   62
         Top             =   2760
         Width           =   4695
      End
      Begin VB.Label xncuenta2 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   5040
         TabIndex        =   61
         Top             =   3240
         Width           =   4695
      End
      Begin VB.Label xncuenta3 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   5040
         TabIndex        =   60
         Top             =   3720
         Width           =   4695
      End
      Begin VB.Label Label34 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero"
         Height          =   495
         Left            =   3240
         TabIndex        =   59
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   14715
      TabIndex        =   27
      Top             =   7680
      Visible         =   0   'False
      Width           =   14775
      Begin VB.CommandButton Command9 
         Caption         =   "MasRapido"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3960
         TabIndex        =   31
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Retorna"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8640
         TabIndex        =   30
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Borra"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7080
         TabIndex        =   29
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Agrega"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5520
         TabIndex        =   28
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   6855
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   13095
      Begin VB.TextBox moneda 
         Height          =   495
         Left            =   6840
         MaxLength       =   1
         TabIndex        =   92
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5040
         TabIndex        =   25
         Top             =   4320
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Grabar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2520
         TabIndex        =   24
         Top             =   4320
         Width           =   2055
      End
      Begin VB.TextBox ccosto 
         Height          =   495
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   23
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox glosa 
         Height          =   495
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   21
         Top             =   1560
         Width           =   5295
      End
      Begin VB.TextBox paridad 
         Height          =   495
         Left            =   9000
         MaxLength       =   10
         TabIndex        =   19
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox cuenta 
         Height          =   495
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox haber 
         Height          =   495
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   15
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox debe 
         Height          =   495
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   14
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label ncuenta 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   3000
         TabIndex        =   32
         Top             =   360
         Width           =   7215
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Centro Costos"
         Height          =   495
         Left            =   240
         TabIndex        =   22
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Glosa"
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Cambio"
         Height          =   495
         Left            =   7680
         TabIndex        =   18
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Moneda"
         Height          =   495
         Left            =   5400
         TabIndex        =   16
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Haber"
         Height          =   495
         Left            =   2760
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Debe"
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuenta"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acepta"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   7
      Top             =   840
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid dbgrid1 
      Height          =   4815
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   8493
      _Version        =   393216
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
   Begin VB.TextBox fecha 
      Height          =   495
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox voucher 
      Height          =   495
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox origen 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label ydolar 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   10680
      TabIndex        =   38
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label ysoles 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   10680
      TabIndex        =   37
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label ydebedolar 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   6960
      TabIndex        =   36
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label yhaberdolar 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   8760
      TabIndex        =   35
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label yhabersoles 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   8760
      TabIndex        =   34
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label ydebesoles 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   6960
      TabIndex        =   33
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label mestrabajo 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   13560
      TabIndex        =   26
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Totales Dolares"
      Height          =   495
      Left            =   4680
      TabIndex        =   9
      Top             =   7080
      Width           =   2295
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Totales Moneda Local"
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nro"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Origen"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Menu lfo444 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tsiscont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bufferx_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    filtro

End Sub

Private Sub ccosto_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(Trim(ccosto)) > 0 Then
        found = busca_ccosto()

    End If

    cuenta.SetFocus

End Sub

Private Sub Command1_Click()

    If Len(Trim(origen)) = 0 Then
        Exit Sub

    End If

    If Len(Trim(voucher)) = 0 Then
        Exit Sub

    End If

    If Not IsDate(fecha) Then
        Exit Sub

    End If

    If Mid$(mestrabajo, 3, 4) <> Mid$(fecha, 7, 4) Then
        Exit Sub

    End If

    sql_cabeza
    Command1.Visible = False
    Picture1.Visible = True
    habilita 1

End Sub

Sub sql_cabeza()

    Dim mytablex   As New ADODB.Recordset

    Dim buf        As String

    Dim debesoles  As Double

    Dim habersoles As Double

    Dim debedolar  As Double

    Dim haberdolar As Double

    buf = "select Voucher.Cuenta,cuentas.Descripcio,Voucher.Debe,Voucher.Haber,Voucher.moneda as M,Voucher.paridad, "
    buf = buf & " Voucher.tipo,Voucher.numero,Voucher.codigo,Voucher.nombre,Voucher.Glosa"
    buf = buf & " from voucher left join cuentas on voucher.cuenta=cuentas.cuenta where voucher.voucher=" & Val(voucher) & " and voucher.origen='" & extra_loquesea(origen) & "'"
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = mytablex
    dbGrid1.columns(0).Width = 1500
    dbGrid1.columns(1).Width = 4000
    dbGrid1.columns(2).Width = 1000
    dbGrid1.columns(3).Width = 1000
    dbGrid1.columns(4).Width = 700
    dbGrid1.columns(5).Width = 1000
    dbGrid1.columns(6).Width = 800
    dbGrid1.columns(7).Width = 1500
    dbGrid1.columns(8).Width = 1500
    dbGrid1.columns(9).Width = 3500
    dbGrid1.columns(10).Width = 3500
    dbGrid1.refresh

    debesoles = 0
    habersoles = 0

    debedolar = 0
    haberdolar = 0
    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("m") = "S" Then
            debesoles = debesoles + Val("" & mytablex.Fields("debe"))
            habersoles = habersoles + Val("" & mytablex.Fields("haber"))

        End If

        If "" & mytablex.Fields("m") = "D" Then
            debedolar = debedolar + Val("" & mytablex.Fields("debe"))
            haberdolar = haberdolar + Val("" & mytablex.Fields("haber"))

        End If

        mytablex.MoveNext
    Loop
    ydebesoles = Format(debesoles, "0.00")
    ydebedolar = Format(debedolar, "0.00")
    yhabersoles = Format(habersoles, "0.00")
    yhaberdolar = Format(haberdolar, "0.00")

    ysoles = Format(debesoles - habersoles, "0.00")
    ydolar = Format(debedolar - haberdolar, "0.00")

End Sub

Private Sub Command10_Click()
    filtro

End Sub

Private Sub Command2_Click()

    Dim found As Integer

    found = valida()

    If found = 0 Then
        MsgBox "Datos invalidos ", 48, "Aviso"
        Exit Sub

    End If

    found = grabar()
    inicializa

End Sub

Function grabar()

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from voucher where voucher=" & Val(voucher) & " and origen='" & extra_loquesea(origen) & "'", cn, adOpenStatic, adLockOptimistic
    'If mytablex.RecordCount > 0 Then
    '   MsgBox "Ya existe numero ", 48, "Aviso"
    '   mytablex.Close
    '   Exit Function
    'End If
    mytablex.AddNew
    mytablex.Fields("voucher") = Val(voucher)
    mytablex.Fields("origen") = extra_loquesea(origen)
    mytablex.Fields("cuenta") = Trim(cuenta)
    mytablex.Fields("debe") = Val(debe)
    mytablex.Fields("haber") = Val(haber)
    mytablex.Fields("ccosto") = Trim(ccosto)
    mytablex.Fields("moneda") = Trim(moneda)
    mytablex.Fields("paridad") = Val(paridad)
    mytablex.Fields("glosa") = Trim(glosa)
    mytablex.Update
    mytablex.Close
    adiciona_numero
    inicializa
    sql_cabeza
    cuenta.SetFocus

End Function

Function valida()

    Dim found As Integer

    If Len(Trim(cuenta)) = 0 Then
        Exit Function

    End If

    If moneda <> "S" And moneda <> "D" Then
        moneda.SetFocus
        Exit Function

    End If

    found = busca_cuenta()

    If found = 0 Then
        cuenta.SetFocus
        Exit Function

    End If

    If Len(Trim(ccosto)) > 0 Then
        found = busca_ccosto()

        If found = 0 Then
            ccosto.SetFocus
            Exit Function

        End If

    End If

    valida = 1

End Function

Function valida1()

    Dim found As Integer

    If xmoneda <> "S" And xmoneda <> "D" Then
        xmoneda.SetFocus
        Exit Function

    End If

    If Len(xcodigo) = 0 Then
        xcodigo.SetFocus
        Exit Function

    End If

    found = busca_xcodigo()

    If found = 0 Then
        xcodigo.SetFocus
        Exit Function

    End If

    If Len(xtipo) = 0 Then
        xtipo.SetFocus
        Exit Function

    End If

    found = busca_xtipo()

    If found = 0 Then
        xtipo.SetFocus
        Exit Function

    End If

    If Len(Trim(xcuenta1)) = 0 Then
        xcuenta1.SetFocus
        Exit Function

    End If
   
    found = busca_xcuenta1()

    If found = 0 Then
        xcuenta1.SetFocus
        Exit Function

    End If

    If Len(xcuenta2) > 0 Then
        found = busca_xcuenta2()

        If found = 0 Then
            xcuenta2.SetFocus
            Exit Function

        End If

    End If

    If Len(xcuenta3) > 0 Then
        found = busca_xcuenta3()

        If found = 0 Then
            xcuenta3.SetFocus
            Exit Function

        End If

    End If

    valida1 = 1

End Function

Function grabar1()

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    mytabley.Open "select * from cuentasparametro", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount = 0 Then
        MsgBox "No hay parametros de cuenta ", 48, "Aviso"
        mytabley.Close
        Exit Function

    End If

    mytablex.Open "select * from voucher where voucher=" & Val(voucher) & " and origen='" & extra_loquesea(origen) & "'", cn, adOpenStatic, adLockOptimistic
    'If mytablex.RecordCount > 0 Then
    '   MsgBox "Ya existe numero ", 48, "Aviso"
    '   mytablex.Close
    '   Exit Function
    'End If

    mytablex.AddNew
    mytablex.Fields("voucher") = Val(voucher)
    mytablex.Fields("origen") = extra_loquesea(origen)
    mytablex.Fields("cuenta") = Trim(xcuenta1)
    mytablex.Fields("debe") = Val(xmonto1)
    mytablex.Fields("haber") = 0 'Val(haber)
    mytablex.Fields("ccosto") = "" 'Trim(ccosto)
    mytablex.Fields("moneda") = Trim(xmoneda)
    mytablex.Fields("paridad") = Val(xparidad)
    mytablex.Fields("glosa") = Trim(xglosa)
    mytablex.Fields("tipo") = Trim(xtipo)
    mytablex.Fields("numero") = Trim(xnumero)
    mytablex.Fields("codigo") = Trim(xcodigo)
    mytablex.Fields("nombre") = Trim(xnombre)
    mytablex.Fields("fecha") = Trim(xfecha)
    mytablex.Fields("fechav") = Trim(xfechav)
    mytablex.Update
   
    If Len(Trim(xcuenta2)) > 0 Then
        mytablex.AddNew
        mytablex.Fields("voucher") = Val(voucher)
        mytablex.Fields("origen") = extra_loquesea(origen)
        mytablex.Fields("cuenta") = Trim(xcuenta2)
        mytablex.Fields("debe") = Val(xmonto2)
        mytablex.Fields("haber") = 0 'Val(haber)
        mytablex.Fields("ccosto") = "" 'Trim(ccosto)
        mytablex.Fields("moneda") = Trim(xmoneda)
        mytablex.Fields("paridad") = Val(xparidad)
        mytablex.Fields("glosa") = Trim(xglosa)
        mytablex.Fields("tipo") = Trim(xtipo)
        mytablex.Fields("numero") = Trim(xnumero)
        mytablex.Fields("codigo") = Trim(xcodigo)
        mytablex.Fields("nombre") = Trim(xnombre)
        mytablex.Fields("fecha") = Trim(xfecha)
        mytablex.Fields("fechav") = Trim(xfechav)
        mytablex.Update

    End If
   
    If Len(Trim(xcuenta3)) > 0 Then
        mytablex.AddNew
        mytablex.Fields("voucher") = Val(voucher)
        mytablex.Fields("origen") = extra_loquesea(origen)
        mytablex.Fields("cuenta") = Trim(xcuenta3)
        mytablex.Fields("debe") = Val(xmonto3)
        mytablex.Fields("haber") = 0 'Val(haber)
        mytablex.Fields("ccosto") = "" 'Trim(ccosto)
        mytablex.Fields("moneda") = Trim(xmoneda)
        mytablex.Fields("paridad") = Val(xparidad)
        mytablex.Fields("glosa") = Trim(xglosa)
        mytablex.Fields("tipo") = Trim(xtipo)
        mytablex.Fields("numero") = Trim(xnumero)
        mytablex.Fields("codigo") = Trim(xcodigo)
        mytablex.Fields("nombre") = Trim(xnombre)
        mytablex.Fields("fecha") = Trim(xfecha)
        mytablex.Fields("fechav") = Trim(xfechav)
        mytablex.Update

    End If

    mytablex.AddNew
    mytablex.Fields("voucher") = Val(voucher)
    mytablex.Fields("origen") = extra_loquesea(origen)

    Select Case extra_loquesea(origen)

        Case "01"   'compras
            mytablex.Fields("cuenta") = Trim("" & mytabley.Fields("igvcompra"))

        Case "02"   'ventas
            mytablex.Fields("cuenta") = Trim("" & mytabley.Fields("igvventa"))

    End Select

    mytablex.Fields("debe") = Val(xigvbasegravada)
    mytablex.Fields("haber") = 0 'Val(haber)
    mytablex.Fields("ccosto") = "" 'Trim(ccosto)
    mytablex.Fields("moneda") = Trim(xmoneda)
    mytablex.Fields("paridad") = Val(xparidad)
    mytablex.Fields("glosa") = Trim(xglosa)
   
    mytablex.Fields("tipo") = Trim(xtipo)
    mytablex.Fields("numero") = Trim(xnumero)
    mytablex.Fields("codigo") = Trim(xcodigo)
    mytablex.Fields("nombre") = Trim(xnombre)
    mytablex.Fields("fecha") = Trim(xfecha)
    mytablex.Fields("fechav") = Trim(xfechav)
    mytablex.Update

    mytablex.AddNew
    mytablex.Fields("voucher") = Val(voucher)
    mytablex.Fields("origen") = extra_loquesea(origen)

    Select Case extra_loquesea(origen)

        Case "01"   'compras

            If xmoneda = "S" Then
                mytablex.Fields("cuenta") = Trim("" & mytabley.Fields("ctacomsoles"))

            End If

            If xmoneda = "D" Then
                mytablex.Fields("cuenta") = Trim("" & mytabley.Fields("ctacomdolar"))

            End If

        Case "02"   'VENTAS

            If xmoneda = "S" Then
                mytablex.Fields("cuenta") = Trim("" & mytabley.Fields("ctacobsoles"))

            End If

            If xmoneda = "D" Then
                mytablex.Fields("cuenta") = Trim("" & mytabley.Fields("ctacobdolar"))

            End If

    End Select

    mytablex.Fields("debe") = 0
    mytablex.Fields("haber") = Val(xigvbasegravada) + Val(xbasegravada)
    mytablex.Fields("ccosto") = "" 'Trim(ccosto)
    mytablex.Fields("moneda") = Trim(xmoneda)
    mytablex.Fields("paridad") = Val(xparidad)
    mytablex.Fields("glosa") = Trim(xglosa)
    mytablex.Fields("tipo") = Trim(xtipo)
    mytablex.Fields("numero") = Trim(xnumero)
    mytablex.Fields("codigo") = Trim(xcodigo)
    mytablex.Fields("nombre") = Trim(xnombre)
    mytablex.Fields("fecha") = Trim(xfecha)
    mytablex.Fields("fechav") = Trim(xfechav)
   
    mytablex.Update

    mytablex.Close
    mytabley.Close
    adiciona_numero
    inicializa1
    sql_cabeza
    xfecha.SetFocus

End Function

Private Sub Command3_Click()
    Frame1.Visible = False

End Sub

Private Sub Command4_Click()

    Dim found As Integer

    found = valida1()

    If found = 0 Then
        MsgBox "Datos invalidos ", 48, "Aviso"
        Exit Sub

    End If

    found = grabar1()
    inicializa1

End Sub

Private Sub Command5_Click()
    Frame2.Visible = False

End Sub

Private Sub Command6_Click()
    Frame1.Visible = True
    inicializa
    refresca_dolares
    cuenta.SetFocus

End Sub

Sub inicializa()
    cuenta = ""
    ncuenta = ""
    haber = ""
    debe = ""
    paridad = ""
    ccosto = ""
    glosa = ""
    nccosto = ""
    moneda = "S"

End Sub

Sub inicializa1()
    xncodigo = ""
    xmoneda = "S"
    xparidad = "1"
    xcuenta1 = ""
    xcuenta2 = ""
    xcuenta3 = ""
    xncuenta1 = ""
    xncuenta2 = ""
    xncuenta3 = ""
    xmonto1 = ""
    xmonto2 = ""
    xmonto3 = ""
    xcuenta1 = ""
    xcuenta2 = ""
    xcuenta3 = ""
    xbasegravada = ""
    xexportacion = ""
    xexonerado = ""
    xinafecto = ""
    xisc = ""
    xigvbasegravada = ""
    xigvexportacion = ""
    xigvexonerado = ""
    xotros = ""
    xcodigo = ""
    xglosa = ""
    xnumero = ""

End Sub

Private Sub Command8_Click()
    Command1.Visible = True
    Picture1.Visible = False
    habilita 0

End Sub

Private Sub Command9_Click()
    Frame2.Visible = True

    inicializa1

    xfecha = fecha
    xfechav = fecha
    refresca_dolares
    xfecha.SetFocus

End Sub

Private Sub cuenta_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(cuenta) > 0 Then
        found = busca_cuenta()

        If found = 0 Then
            cuenta = ""

        End If

    End If

    debe.SetFocus

End Sub

Private Sub dbgrid2_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        bufferx.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "22" Then
            xtipo = Trim(DBGrid2.columns(1))
            Frame4.Visible = False
            'Frame4.Enabled = False
            xtipo.SetFocus
            Exit Sub

        End If

        If opcion1 = "23" Then
            xcodigo = Trim(DBGrid2.columns(1))
            Frame4.Visible = False
            'Frame4.Enabled = False
            xcodigo.SetFocus
            Exit Sub

        End If

        If opcion1 = "24" Then
            xcuenta1 = Trim(DBGrid2.columns(0))
            Frame4.Visible = False
            'Frame4.Enabled = False
            xcuenta1.SetFocus
            Exit Sub

        End If

    End If

End Sub

Private Sub debe_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    haber.SetFocus

End Sub

Private Sub Form_Load()
    pone_mestrabajo
    carga_origen
    fecha = Format(Now, "dd/mm/yyyy")

End Sub

Sub calcula()

    Dim sdx  As Double

    Dim sdx1 As Double

    sdx = Val(xmonto1) + Val(xmonto2) + Val(xmonto3)
    sdx = Val(Format(sdx, "0.00"))
    xbasegravada = "" & sdx
    sdx1 = sdx * 18 / 100
    sdx1 = Val(Format(sdx1, "0.00"))
    xigvbasegravada = "" & sdx1

End Sub

Private Sub glosa_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    ccosto.SetFocus

End Sub

Private Sub haber_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    moneda.SetFocus

End Sub

Private Sub Label17_Click()
    menu_tipodoc

End Sub

Private Sub Label18_Click()
    menu_codigo

End Sub

Private Sub lfo444_Click()

    If Frame4.Visible = True Then
        Frame4.Visible = False
        Exit Sub

    End If

    tsiscont.Hide
    Unload tsiscont

End Sub

Sub carga_origen()

    Dim mytablex As New ADODB.Recordset

    origen.Clear
    origen.AddItem ""
    mytablex.Open "select * from origen ordeR by origen", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        origen.AddItem Trim("" & mytablex.Fields("origen")) & "|" & Trim("" & mytablex.Fields("descripcio"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    origen.ListIndex = 0

End Sub

Sub pone_mestrabajo()

    Dim mytablex As New ADODB.Recordset

    mestrabajo = ""
    mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        mestrabajo = Trim("" & mytablex.Fields("mesconta")) & Trim("" & mytablex.Fields("anoconta"))

    End If

    mytablex.Close

End Sub

Sub pone_numero()

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    voucher = ""
    mytablex.Open "select * from origen where origen='" & extra_loquesea(Trim(origen)) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        sdx = Val("" & mytablex.Fields("numero" & Mid$(mestrabajo, 1, 2))) + 1
        voucher = "" & sdx

    End If

    mytablex.Close

End Sub

Function busca_cuenta()

    Dim mytablex As New ADODB.Recordset

    ncuenta = ""
    mytablex.Open "select * from cuentas where cuenta='" & Trim(cuenta) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        ncuenta = Trim("" & mytablex.Fields("descripcio"))
        busca_cuenta = 1

    End If

    mytablex.Close

End Function

Function busca_ccosto()

    Dim mytablex As New ADODB.Recordset

    nccosto = ""
    mytablex.Open "select * from ccosto where ccosto='" & Trim(ccosto) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        ccosto = Trim("" & mytablex.Fields("descripcio"))
        busca_ccosto = 1

    End If

    mytablex.Close

End Function

Private Sub moneda_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    paridad.SetFocus

End Sub

Private Sub origen_Click()
    pone_numero
    sql_cabeza

End Sub

Private Sub paridad_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    glosa.SetFocus

End Sub

Sub adiciona_numero()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from origen where origen='" & extra_loquesea(Trim(origen)) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        mytablex.Fields("numero" & Mid$(mestrabajo, 1, 2)) = Trim(voucher)
        mytablex.Update

    End If

    mytablex.Close

End Sub

Sub habilita(sw As Integer)

    If sw = 0 Then
        xsw = True

    End If

    If sw = 1 Then
        xsw = False

    End If

    origen.Enabled = xsw
    voucher.Enabled = xsw
    fecha.Enabled = xsw

End Sub

Private Sub xbasegravada_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    xexonerado.SetFocus

End Sub

Private Sub xcodigo_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(Trim(xcodigo)) = 0 Then
        menu_codigo
        Exit Sub

    End If

    If Len(xcodigo) > 0 Then
        found = busca_xcodigo()

        If found = 0 Then
            xncodigo = ""

        End If

    End If

    xglosa.SetFocus

End Sub

Private Sub xcuenta1_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If xmoneda <> "S" And xmoneda <> "D" Then
        xmoneda.SetFocus
        Exit Sub

    End If

    If KeyAscii <> 13 Then Exit Sub
    If Len(Trim(xcuenta1)) = 0 Then
        menu_xcuenta
        Exit Sub

    End If

    If Len(xcuenta1) > 0 Then
        found = busca_xcuenta1()

        If found = 0 Then
            xcuenta1 = ""

        End If

    End If

    xmonto2.SetFocus

End Sub

Private Sub xcuenta2_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If xmoneda <> "S" And xmoneda <> "D" Then
        xmoneda.SetFocus
        Exit Sub

    End If

    If KeyAscii <> 13 Then Exit Sub
    If Len(xcuenta2) > 0 Then
        found = busca_xcuenta2()

        If found = 0 Then
            xcuenta2 = ""

        End If

    End If

    xmonto3.SetFocus

End Sub

Private Sub xcuenta3_KeyPress(KeyAscii As Integer)

    If xmoneda <> "S" And xmoneda <> "D" Then
        xmoneda.SetFocus
        Exit Sub

    End If

    If KeyAscii <> 13 Then Exit Sub
    If Len(xcuenta3) > 0 Then
        found = busca_xcuenta3()

        If found = 0 Then
            xcuenta3 = ""

        End If

    End If

    xbasegravada.SetFocus

End Sub

Private Sub xexonerado_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    xinafecto.SetFocus

End Sub

Private Sub xfecha_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    xfechav.SetFocus

End Sub

Private Sub xfechav_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    xtipo.SetFocus

End Sub

Private Sub xglosa_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    xmonto1.SetFocus

End Sub

Private Sub xinafecto_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    xisc.SetFocus

End Sub

Private Sub xisc_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    xotros.SetFocus

End Sub

Private Sub xmonto1_KeyPress(KeyAscii As Integer)

    If xmoneda <> "S" And xmoneda <> "D" Then
        xmoneda.SetFocus
        Exit Sub

    End If

    If KeyAscii <> 13 Then Exit Sub
    calcula
    xcuenta1.SetFocus

End Sub

Private Sub xmonto2_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    calcula
    xcuenta2.SetFocus

End Sub

Private Sub xmonto3_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    calcula
    xcuenta3.SetFocus

End Sub

Private Sub xnumero_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    xcodigo.SetFocus

End Sub

Private Sub xotros_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    xfecha.SetFocus

End Sub

Private Sub xtipo_keyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(Trim(xtipo)) = 0 Then
        menu_tipodoc
        Exit Sub

    End If

    xnumero.SetFocus

End Sub

Function busca_xcuenta1()

    Dim mytablex As New ADODB.Recordset

    xncuenta1 = ""
    mytablex.Open "select * from cuentas where cuenta='" & Trim(xcuenta1) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        xncuenta1 = Trim("" & mytablex.Fields("descripcio"))
        busca_xcuenta1 = 1

    End If

    mytablex.Close

End Function

Function busca_xcuenta2()

    Dim mytablex As New ADODB.Recordset

    xncuenta2 = ""
    mytablex.Open "select * from cuentas where cuenta='" & Trim(xcuenta2) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        xncuenta2 = Trim("" & mytablex.Fields("descripcio"))
        busca_xcuenta2 = 1

    End If

    mytablex.Close

End Function

Function busca_xcuenta3()

    Dim mytablex As New ADODB.Recordset

    xncuenta3 = ""
    mytablex.Open "select * from cuentas where cuenta='" & Trim(xcuenta3) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        xncuenta3 = Trim("" & mytablex.Fields("descripcio"))
        busca_xcuenta3 = 1

    End If

    mytablex.Close

End Function

Function busca_xtipo()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from docta where docta='" & Trim(xtipo) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_xtipo = 1

    End If

    mytablex.Close

End Function

Function busca_xcodigo()

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    buf = ""

    If extra_loquesea(origen) = "01" Then  'compras
        buf = "proveedo"

    End If

    If extra_loquesea(origen) = "02" Then  'ventas
        buf = "clientes"

    End If

    If Len(buf) = 0 Then Exit Function
    xncodigo = ""
    mytablex.Open "select * from " & buf & " where codigo='" & Trim(xcodigo) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        xncodigo = Trim("" & mytablex.Fields("nombre"))
        busca_xcodigo = 1

    End If

    mytablex.Close

End Function

Sub refresca_dolares()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from tcambio where fecha='" & fecha & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        xparidad = Format(Val("" & mytablex.Fields("venta")), "0.000")

    End If

    mytablex.Close

End Sub

Sub menu_tipodoc()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.ListIndex = 0
    opcion1 = "22"
    bufferx = ""
    Frame4.Visible = True
    filtro

End Sub

Sub menu_codigo()
    Combo2.Clear
    Combo2.AddItem "Nombre"
    Combo2.ListIndex = 0
    opcion1 = "23"
    bufferx = ""
    Frame4.Visible = True
    filtro

End Sub

Sub menu_xcuenta()
    Combo2.Clear
    Combo2.AddItem "Cuenta"
    Combo2.ListIndex = 0
    opcion1 = "24"
    bufferx = ""
    Frame4.Visible = True
    filtro

End Sub

Sub filtro()

    Dim mytablex As New ADODB.Recordset

    Dim cad      As String

    If opcion1 = "22" Then  'tipo documento
        If Len(bufferx) = 0 Then
            cad = "select Descripcio,docta from docta "

        End If

        If Len(bufferx) > 0 Then
            cad = "select Descripcio,docta from docta where " & Combo2 & " like '" & bufferx & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set DBGrid2.DataSource = mytablex
        DBGrid2.columns(0).Width = 5000
        DBGrid2.columns(1).Width = 2000

    End If

    If opcion1 = "23" Then  'tipo documento
        If Len(bufferx) = 0 Then
            cad = "select Nombre,Codigo from clientes "

        End If

        If Len(bufferx) > 0 Then
            cad = "select Nombre,Codigo from clientes where " & Combo2 & " like '" & bufferx & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set DBGrid2.DataSource = mytablex
        DBGrid2.columns(0).Width = 5000
        DBGrid2.columns(1).Width = 2000

    End If
   
    If opcion1 = "24" Then  'cuenta
        If Len(bufferx) = 0 Then
            cad = "select Cuenta,Descripcio from cuentas order by cuenta "

        End If

        If Len(bufferx) > 0 Then
            cad = "select Cuenta,Descripcio from cuentas where " & Combo2 & " like '" & bufferx & "%' order by cuenta"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set DBGrid2.DataSource = mytablex
        DBGrid2.columns(1).Width = 5000
        DBGrid2.columns(0).Width = 2000

    End If
   
    If mytablex.RecordCount > 0 Then
        DBGrid2.SetFocus

    End If

    Exit Sub

End Sub

