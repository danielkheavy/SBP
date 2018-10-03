VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form tparame 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Tabla de Parametros Generales"
   ClientHeight    =   8535
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   12030
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   0
      TabIndex        =   70
      Top             =   120
      Visible         =   0   'False
      Width           =   11895
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
         Left            =   8280
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   600
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   600
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
         Height          =   495
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   600
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   6735
         Left            =   240
         TabIndex        =   74
         Top             =   1200
         Width           =   9735
         _ExtentX        =   17171
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
   Begin VB.TextBox tipo5 
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
      TabIndex        =   68
      Top             =   7200
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Height          =   6735
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox aduana 
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
         TabIndex        =   66
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox prehora 
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
         TabIndex        =   64
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox banco 
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
         TabIndex        =   62
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox conteo 
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
         MaxLength       =   11
         TabIndex        =   60
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox plocal 
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
         TabIndex        =   58
         Top             =   3000
         Width           =   255
      End
      Begin VB.TextBox vdolar 
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
         TabIndex        =   56
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox centraliza 
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
         TabIndex        =   54
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox insumo 
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
         TabIndex        =   52
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox proveedo 
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
         MaxLength       =   11
         TabIndex        =   50
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox parivta 
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
         TabIndex        =   48
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox paricomp 
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
         TabIndex        =   46
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox clientes 
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
         MaxLength       =   11
         TabIndex        =   44
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label31 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aduana"
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
         Left            =   120
         TabIndex        =   67
         Top             =   4440
         Width           =   2175
      End
      Begin VB.Label Label30 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Prehora"
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
         Left            =   120
         TabIndex        =   65
         Top             =   4080
         Width           =   2175
      End
      Begin VB.Label Label29 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Banco"
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
         Left            =   120
         TabIndex        =   63
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label Label28 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Conteo"
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
         Left            =   120
         TabIndex        =   61
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label Label27 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "pLocal"
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
         Left            =   120
         TabIndex        =   59
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ve Dolar"
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
         Left            =   120
         TabIndex        =   57
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Centraliza"
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
         Left            =   120
         TabIndex        =   55
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Insumo"
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
         Left            =   120
         TabIndex        =   53
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proveedor"
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
         Left            =   120
         TabIndex        =   51
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ventas T/C"
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
         Left            =   120
         TabIndex        =   49
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Compras T/C"
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
         Left            =   120
         TabIndex        =   47
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clientes"
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
         Left            =   120
         TabIndex        =   45
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.TextBox ocurrencia 
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
      Left            =   5760
      MaxLength       =   10
      TabIndex        =   41
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox pocket 
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
      TabIndex        =   39
      Top             =   6840
      Width           =   2295
   End
   Begin VB.TextBox SERVIDOR 
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
      TabIndex        =   37
      Top             =   6480
      Width           =   495
   End
   Begin VB.TextBox tradiario 
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
      TabIndex        =   35
      Top             =   6120
      Width           =   495
   End
   Begin VB.TextBox imp_und 
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
      TabIndex        =   33
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox pedauto 
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
      TabIndex        =   31
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox deliveri 
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
      MaxLength       =   3
      TabIndex        =   29
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox cabecera2 
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
      MaxLength       =   60
      TabIndex        =   28
      Top             =   5400
      Width           =   5775
   End
   Begin VB.TextBox cabecera1 
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
      MaxLength       =   60
      TabIndex        =   26
      Top             =   5040
      Width           =   5775
   End
   Begin VB.TextBox produccion 
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
      MaxLength       =   11
      TabIndex        =   24
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox anoconta 
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
      TabIndex        =   22
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox mesconta 
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
      MaxLength       =   2
      TabIndex        =   20
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox saldoini 
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
      TabIndex        =   17
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox bodega 
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
      MaxLength       =   2
      TabIndex        =   15
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox igv 
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
      TabIndex        =   13
      Top             =   1920
      Width           =   1335
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
      MaxLength       =   10
      TabIndex        =   11
      Top             =   1560
      Width           =   1335
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
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1200
      Width           =   5775
   End
   Begin VB.TextBox codigo 
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
      MaxLength       =   2
      TabIndex        =   0
      Top             =   840
      Width           =   1335
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
      Left            =   1440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tparame.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Grabar registro"
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
      Picture         =   "tparame.frx":1212
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Nuevo registro"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdHelp 
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
      Left            =   3600
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tparame.frx":2424
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Ayuda"
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
      Picture         =   "tparame.frx":3636
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir"
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
      Left            =   4320
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tparame.frx":4848
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
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
      Left            =   720
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tparame.frx":5A5A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Borrar registro"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tparame.frx":6C6C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Consulta"
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label26 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Visualiza Tipo 5"
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
      Left            =   120
      TabIndex        =   69
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OcurrenciasGrabadas"
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
      Left            =   3360
      TabIndex        =   42
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CorrelativoTicketIng"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   40
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Servidor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   38
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Transaccion (D)iario"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   36
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Veprodc.Cuadre(1=s)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   34
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TipoDoc.Ped.Autom"
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
      Left            =   120
      TabIndex        =   32
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TipoDoc.Deliveri"
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
      Left            =   120
      TabIndex        =   30
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cabecera_Reporte"
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
      Left            =   120
      TabIndex        =   27
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Orden Produccion"
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
      Left            =   120
      TabIndex        =   25
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Año Contable"
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
      Left            =   120
      TabIndex        =   23
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mes Contable"
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
      Left            =   120
      TabIndex        =   21
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DD/MM/YYYY"
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
      Left            =   4320
      TabIndex        =   19
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo Inicial"
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
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bodega"
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
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Igv"
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
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Producto"
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
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripcion"
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
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo"
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
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   2175
   End
   Begin VB.Menu ajdu1 
      Caption         =   "&Add"
   End
   Begin VB.Menu bo712 
      Caption         =   "&Borrar"
   End
   Begin VB.Menu grba1 
      Caption         =   "&Grabar"
   End
   Begin VB.Menu djuer1 
      Caption         =   "&Reporte"
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tparame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ajdu1_Click()
If Frame1.Visible = True Then Exit Sub
inicializa
codigo = ""
codigo.SetFocus

End Sub

Private Sub bo712_Click()
Dim found As Integer
If Frame1.Visible = True Then Exit Sub
found = busca_registro()
If found = 0 Then
   MsgBox "No existe registro", 48, "Aviso"
   Exit Sub
End If
found = borra_registro()
If found = 0 Then Exit Sub
MsgBox "Ok,Registro Borrado", 48, "Aviso"
codigo = ""
inicializa
codigo.SetFocus
End Sub

Private Sub bodega_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
saldoini.SetFocus

End Sub

Private Sub buffer_DblClick()
Command1_Click
End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   Frame1.Visible = False
   Frame1.Enabled = False
   codigo.SetFocus
   Exit Sub
End If
Command1_Click

End Sub

Private Sub cmdAddEntry_Click()
ajdu1_Click
End Sub

Private Sub cmdDelete_Click()
bo712_Click
End Sub

Private Sub cmdExit_Click()
dlo132_Click

End Sub

Private Sub cmdPrint_Click()
djuer1_Click
End Sub

Private Sub cmdSave_Click()
grba1_Click
End Sub

Private Sub cmdSort_Click()
Dim cad As String
   Dim rconsulta As New ADODB.Recordset
   cad = "SELECT * FROM parame  "
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True Or rconsulta.BOF = True Then
      rconsulta.Close
      Exit Sub
   End If
opcion1 = "1"
Frame1.Enabled = True
Frame1.Visible = True
buffer = ""
buffer.SetFocus
Command1_Click
End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   dlo132_Click
   Exit Sub
End If
If Len(codigo) = 0 Then Exit Sub
found = busca_registro()
If found = 0 Then
   inicializa
End If
descripcio.SetFocus
End Sub

Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
cmdSort_Click
End If
End Sub

Private Sub Command1_Click()
ejecuta 1
End Sub
Sub ejecuta(sw As Integer)
Dim rconsulta As New ADODB.Recordset
Dim cad As String
If opcion1 = "1" Then  'bodega
   If Len(buffer) = 0 Then
      cad = "SELECT Descripcio,Codigo from parame    "
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT Descripcio,Codigo from parame   where  " & Combo1 & " like '" & buffer & "%'"
   End If
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      rconsulta.Close
      buffer.SetFocus
      Exit Sub
   End If
   Set dbGrid1.DataSource = rconsulta
   dbGrid1.columns(0).Width = 4000
   dbGrid1.columns(1).Width = 2000
   
   
   If sw = 1 Then
      dbGrid1.SetFocus
   End If
   Exit Sub
End If

End Sub



Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
   codigo = dbGrid1.columns(1)
   Frame1.Visible = False
   Frame1.Enabled = False
   codigo.SetFocus
   codigo_KeyPress 13
End If
End Sub


Private Sub descripcio_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
producto.SetFocus

End Sub



Private Sub descripcio_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   codigo.SetFocus
   Exit Sub
End If

End Sub

Private Sub djuer1_Click()
If Frame1.Visible = True Then Exit Sub
reporgen.NAMETABLA = "parame"
reporgen.Show 1

End Sub

Private Sub dlo132_Click()
If Frame1.Visible = True Then
   Frame1.Visible = False
   Exit Sub
End If
tparame.Hide
Unload tparame
End Sub



Private Sub Form_Load()

Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "Codigo"
Combo1.ListIndex = 0
End Sub
Sub inicializa()
clientes = ""
paricomp = ""
parivta = ""
proveedo = ""
insumo = ""
centraliza = ""
vdolar = ""
tipo5 = ""
plocal = ""
conteo = ""
banco = ""
prehora = ""
aduana = ""

ocurrencia = ""
pocket = ""
SERVIDOR = ""
tradiario = ""
imp_und = ""
pedauto = ""
deliveri = ""
cabecera1 = ""
cabecera2 = ""
produccion = ""
descripcio = ""
igv = ""
producto = ""
bodega = ""
saldoini = ""
mesconta = ""
anoconta = ""
End Sub
Function borra_registro()
On Error GoTo cmd56_err
cn.Execute ("DELETE   FROM parame WHERE codigo='" & Trim(codigo) & "'")
borra_registro = 1
Exit Function
cmd56_err:
MsgBox "Aviso en borra " + error$, 48, "Aviso"
Exit Function

End Function
Function busca_registro()
Dim rsexiste As New ADODB.Recordset
   rsexiste.Open "SELECT * FROM parame where codigo='" & Trim(codigo) & "'", cn, adOpenKeyset, adLockOptimistic
   If rsexiste.RecordCount > 0 Then  'si existe
      pone_registro rsexiste
      busca_registro = 1
   End If
End Function
Sub pone_registro(mytablex As ADODB.Recordset)
clientes = "" & mytablex.Fields("clientes")
paricomp = "" & mytablex.Fields("paricomp")
parivta = "" & mytablex.Fields("parivta")
proveedo = "" & mytablex.Fields("proveedo")
insumo = "" & mytablex.Fields("insumo")
centraliza = "" & mytablex.Fields("centraliza")
vdolar = "" & mytablex.Fields("vdolar")
tipo5 = "" & mytablex.Fields("tipo5")
plocal = "" & mytablex.Fields("plocal")
conteo = "" & mytablex.Fields("conteo")
banco = "" & mytablex.Fields("banco")
prehora = "" & mytablex.Fields("prehora")
aduana = "" & mytablex.Fields("aduana")

ocurrencia = "" & mytablex.Fields("ocurrencia")
pocket = "" & mytablex.Fields("pocket")
tradiario = "" & mytablex.Fields("tradiario")
imp_und = "" & mytablex.Fields("imp_und")

pedauto = "" & mytablex.Fields("pedauto")
deliveri = "" & mytablex.Fields("deliveri")
cabecera1 = "" & mytablex.Fields("cabecera1")
cabecera2 = "" & mytablex.Fields("cabecera2")
produccion = "" & mytablex.Fields("produccion")
codigo = "" & mytablex.Fields("codigo")
descripcio = "" & mytablex.Fields("descripcio")
igv = "" & mytablex.Fields("igv")
producto = "" & mytablex.Fields("producto")
bodega = "" & mytablex.Fields("bodega")
saldoini = "" & mytablex.Fields("saldoini")
mesconta = "" & mytablex.Fields("mesconta")
anoconta = "" & mytablex.Fields("anoconta")
End Sub
Sub grabando(sw As Integer)

Dim cad As String


If sw = 0 Then
   cad = "INSERT INTO parame VALUES('" & Trim(codigo) & "',"
   cad = cad & Val(igv) & ",'"
   cad = cad & Trim(producto) & "','"
   cad = cad & Trim(descripcio) & "','"
   cad = cad & Trim(bodega) & "','"
   cad = cad & Trim(saldoini) & "','"
   cad = cad & Trim(mesconta) & "','"
   cad = cad & Trim(anoconta) & "','"
   cad = cad & Trim(produccion) & "','"
   cad = cad & Trim(cabecera1) & "','"
   cad = cad & Trim(cabecera2) & "','"
   cad = cad & Trim(deliveri) & "','"
   cad = cad & Trim(clientes) & "',"
   cad = cad & Val(paricomp) & ","
   cad = cad & Val(parivta) & ",'"
   cad = cad & Trim(proveedo) & "','"
   cad = cad & Trim(insumo) & "','"
   cad = cad & Trim(pedauto) & "','"
   cad = cad & Trim(imp_und) & "','"
   cad = cad & Trim(tradiario) & "','"
   cad = cad & Trim(centraliza) & "','"
   cad = cad & Trim(vdolar) & "','"
   cad = cad & Trim(tipo5) & "','"
   
   cad = cad & Trim(plocal) & "','"
   cad = cad & Trim(conteo) & "','"
   cad = cad & Trim(banco) & "','"
   cad = cad & Trim(prehora) & "',"
   cad = cad & Val(pocket) & ","
   cad = cad & Val(ocurrencia) & ",'"
   cad = cad & Trim(aduana) & "')"
   'MsgBox cad
   cn.Execute (cad)
   MsgBox "Adicion exitosa", 48, "Aviso"
   codigo.SetFocus
   Exit Sub
End If

If sw = 1 Then
   cad = "UPDATE parame SET "
   cad = cad & "igv = " & Val(codigo) & ""
   cad = cad & ",producto = '" & Trim(producto) & "'"
   cad = cad & ",descripcio = '" & Trim(descripcio) & "'"
   cad = cad & ",bodega = '" & Trim(bodega) & "'"
   cad = cad & ",saldoini = '" & Trim(saldoini) & "'"
   cad = cad & ",mesconta = '" & Trim(mesconta) & "'"
   cad = cad & ",anoconta = '" & Trim(anoconta) & "'"
   cad = cad & ",produccion = '" & Trim(produccion) & "'"
   cad = cad & ",cabecera1 = '" & Trim(cabecera1) & "'"
   cad = cad & ",cabecera2 = '" & Trim(cabecera2) & "'"
   cad = cad & ",deliveri = '" & Trim(deliveri) & "'"
   cad = cad & ",clientes = '" & Trim(clientes) & "'"
   cad = cad & ",paricomp = " & Val(paricomp) & ""
   cad = cad & ",proveedo = '" & Trim(proveedo) & "'"
   cad = cad & ",insumo = '" & Trim(insumo) & "'"
   cad = cad & ",pedauto = '" & Trim(pedauto) & "'"
   cad = cad & ",imp_und = '" & Trim(imp_und) & "'"
   cad = cad & ",tradiario = '" & Trim(tradiario) & "'"
   cad = cad & ",centraliza = '" & Trim(centraliza) & "'"
   cad = cad & ",vdolar = '" & Trim(vdolar) & "'"
   cad = cad & ",tipo5 = '" & Trim(tipo5) & "'"
   cad = cad & ",plocal = '" & Trim(plocal) & "'"
   cad = cad & ",conteo = '" & Trim(conteo) & "'"
   cad = cad & ",banco = '" & Trim(banco) & "'"
   cad = cad & ",prehora = '" & Trim(prehora) & "'"
   cad = cad & ",pocket = " & Val(pocket) & ""
   cad = cad & ",ocurrencia = " & Val(ocurrencia) & ""
   cad = cad & ",aduana = '" & Trim(aduana) & "'"
   cad = cad & " WHERE  codigo='" & Trim(codigo) & "'"
   cn.Execute (cad)
   MsgBox "Rescripcion exitosa", 48, "Aviso"
   codigo.SetFocus
   Exit Sub
End If


End Sub

Private Sub grba1_Click()
Dim found As Integer
If Frame1.Visible = True Then Exit Sub
found = grabar()
If found = 0 Then Exit Sub
codigo.SetFocus

End Sub

Private Sub igv_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
bodega.SetFocus

End Sub

Private Sub Label1_Click()
cmdSort_Click
End Sub


Function grabar()
Dim found As Integer
Dim rsexiste As New ADODB.Recordset
found = valida()
If found = 0 Then
   MsgBox "Campos invalidos", 48, "Aviso"
   Exit Function
End If

rsexiste.Open "SELECT * FROM parame where  codigo='" & Trim(codigo) & "'", cn, adOpenKeyset, adLockOptimistic
   If rsexiste.RecordCount > 0 Then  'si existe
      If MsgBox("Desea Reescribir? ", 1, "Aviso") <> 1 Then
         codigo.SetFocus
         Exit Function
      End If
      grabando 1
      Exit Function
   End If
   If MsgBox("Desea Adicionar ? ", 1, "Aviso") <> 1 Then
      codigo.SetFocus
      Exit Function
   End If
   grabando 0
End Function

Function valida()
If Len(codigo) = 0 Then
   codigo.SetFocus
   Exit Function
End If
If Len(descripcio) = 0 Then
   descripcio.SetFocus
   Exit Function
End If
valida = 1
End Function

Private Sub producto_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
igv.SetFocus
End Sub
