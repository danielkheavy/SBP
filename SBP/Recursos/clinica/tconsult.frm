VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tconsult 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla de Consultas"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   13875
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consulta"
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
      Left            =   15
      TabIndex        =   73
      Top             =   -15
      Visible         =   0   'False
      Width           =   12375
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
         TabIndex        =   76
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
         TabIndex        =   75
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
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   600
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid dbGrid2 
         Height          =   6855
         Left            =   240
         TabIndex        =   77
         Top             =   1200
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   12091
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Diagnostico"
      Height          =   8175
      Left            =   1965
      TabIndex        =   58
      Top             =   4035
      Visible         =   0   'False
      Width           =   12375
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Height          =   4695
         Left            =   1920
         TabIndex        =   60
         Top             =   2040
         Visible         =   0   'False
         Width           =   9615
         Begin VB.TextBox enfermedad 
            Height          =   495
            Left            =   1680
            MaxLength       =   11
            TabIndex        =   68
            Top             =   1320
            Width           =   1575
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
            Left            =   8280
            MaskColor       =   &H00FFFFFF&
            Picture         =   "tconsult.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   1080
            UseMaskColor    =   -1  'True
            Width           =   1215
         End
         Begin VB.CommandButton Command5 
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
            Left            =   8280
            MaskColor       =   &H00FFFFFF&
            Picture         =   "tconsult.frx":07AE
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   1215
         End
         Begin VB.TextBox observa1 
            Height          =   495
            Left            =   1680
            MaxLength       =   60
            TabIndex        =   65
            Top             =   1800
            Width           =   6255
         End
         Begin VB.TextBox observa2 
            Height          =   495
            Left            =   1680
            MaxLength       =   60
            TabIndex        =   64
            Top             =   2280
            Width           =   6255
         End
         Begin VB.TextBox observa3 
            Height          =   495
            Left            =   1680
            MaxLength       =   60
            TabIndex        =   63
            Top             =   2760
            Width           =   6255
         End
         Begin VB.TextBox observa4 
            Height          =   495
            Left            =   1680
            MaxLength       =   60
            TabIndex        =   62
            Top             =   3240
            Width           =   6255
         End
         Begin VB.TextBox diagnostico 
            Enabled         =   0   'False
            Height          =   495
            Left            =   1680
            MaxLength       =   11
            TabIndex        =   61
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label25 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Enfermedad"
            Height          =   495
            Left            =   120
            TabIndex        =   72
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label nenfermedad 
            BorderStyle     =   1  'Fixed Single
            Height          =   495
            Left            =   3240
            TabIndex        =   71
            Top             =   1320
            Width           =   4695
         End
         Begin VB.Label Label23 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Comentario"
            Height          =   495
            Left            =   120
            TabIndex        =   70
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label22 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nro"
            Height          =   495
            Left            =   120
            TabIndex        =   69
            Top             =   840
            Width           =   1575
         End
      End
      Begin MSDataGridLib.DataGrid dbGrid3 
         Height          =   4695
         Left            =   120
         TabIndex        =   59
         Top             =   960
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   8281
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
      Begin VB.Label dconsulta 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1080
         TabIndex        =   79
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label dsede 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   78
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   8415
      Left            =   1155
      TabIndex        =   8
      Top             =   4290
      Visible         =   0   'False
      Width           =   12375
      Begin VB.TextBox medicoosi 
         Height          =   375
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   45
         Top             =   6600
         Width           =   1575
      End
      Begin VB.TextBox tipoconsu 
         Height          =   375
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   43
         Top             =   6240
         Width           =   1575
      End
      Begin VB.TextBox fecha 
         Height          =   375
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   41
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox coaseguro 
         Height          =   375
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   39
         Top             =   5880
         Width           =   1575
      End
      Begin VB.TextBox deducible 
         Height          =   375
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   37
         Top             =   5520
         Width           =   1575
      End
      Begin VB.TextBox codigoauto 
         Height          =   375
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   35
         Top             =   5160
         Width           =   1575
      End
      Begin VB.TextBox tipoautor 
         Height          =   375
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   33
         Top             =   4800
         Width           =   1575
      End
      Begin VB.TextBox titularse 
         Height          =   375
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   31
         Top             =   4440
         Width           =   1575
      End
      Begin VB.TextBox tipoafilia 
         Height          =   375
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   29
         Top             =   4080
         Width           =   1575
      End
      Begin VB.TextBox empresa 
         Height          =   375
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   27
         Top             =   3720
         Width           =   1575
      End
      Begin VB.TextBox tiposeguro 
         Height          =   375
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   25
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox medico 
         Height          =   375
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   23
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox clinica 
         Height          =   375
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   21
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox referencia 
         Height          =   375
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   19
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox cliente 
         Height          =   375
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   17
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox sede 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   15
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
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
         Left            =   8040
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tconsult.frx":0F5C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
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
         Left            =   8040
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tconsult.frx":170A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1080
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox codigo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   10
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label empresaldo 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   120
         TabIndex        =   89
         Top             =   7320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label partisaldo 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   1680
         TabIndex        =   88
         Top             =   7320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label nro 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   80
         Top             =   6960
         Width           =   105
      End
      Begin VB.Label ntipoconsu 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3600
         TabIndex        =   57
         Top             =   6240
         Width           =   4215
      End
      Begin VB.Label nmedicoosi 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3600
         TabIndex        =   56
         Top             =   6600
         Width           =   4215
      End
      Begin VB.Label nsede 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3600
         TabIndex        =   55
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label ntipoautor 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3600
         TabIndex        =   54
         Top             =   4800
         Width           =   4215
      End
      Begin VB.Label ntitularse 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3600
         TabIndex        =   53
         Top             =   4440
         Width           =   4215
      End
      Begin VB.Label ntipoafilia 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3600
         TabIndex        =   52
         Top             =   4080
         Width           =   4215
      End
      Begin VB.Label nempresa 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3600
         TabIndex        =   51
         Top             =   3720
         Width           =   4215
      End
      Begin VB.Label ntiposeguro 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3600
         TabIndex        =   50
         Top             =   3360
         Width           =   4215
      End
      Begin VB.Label nmedico 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3600
         TabIndex        =   49
         Top             =   3000
         Width           =   4215
      End
      Begin VB.Label nclinica 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3600
         TabIndex        =   48
         Top             =   2640
         Width           =   4215
      End
      Begin VB.Label nreferencia 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3600
         TabIndex        =   47
         Top             =   2280
         Width           =   4215
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Medico Osi"
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   6600
         Width           =   1575
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Consulta"
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   6240
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Consulta"
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label15 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PagaSeguro"
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   5880
         Width           =   1575
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PagaParticular"
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo Autorizacion"
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Autorizacion"
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Titular Seguro"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Afiliado"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Empresa"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Seguro"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Medico"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clinica"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Referencia"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cliente"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sede"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Nombre 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3600
         TabIndex        =   11
         Top             =   1080
         Width           =   4215
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   13815
      TabIndex        =   3
      Top             =   0
      Width           =   13875
      Begin VB.TextBox fechaf 
         Height          =   375
         Left            =   7440
         MaxLength       =   10
         TabIndex        =   86
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox fechai 
         Height          =   375
         Left            =   7440
         MaxLength       =   10
         TabIndex        =   84
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox xcliente 
         Height          =   375
         Left            =   5040
         MaxLength       =   11
         TabIndex        =   1
         Text            =   "*"
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
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
         Left            =   11040
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tconsult.frx":1EB8
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox xsede 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   5040
         MaxLength       =   6
         TabIndex        =   2
         Top             =   120
         Width           =   1215
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
         Left            =   1440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tconsult.frx":2666
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Ayuda"
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
         Picture         =   "tconsult.frx":3878
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
         Picture         =   "tconsult.frx":4A8A
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
         Picture         =   "tconsult.frx":5C9C
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
         Picture         =   "tconsult.frx":6EAE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Final"
         Height          =   375
         Left            =   6240
         TabIndex        =   87
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   375
         Left            =   6240
         TabIndex        =   85
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cliente"
         Height          =   375
         Left            =   3720
         TabIndex        =   83
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sede"
         Height          =   375
         Left            =   3720
         TabIndex        =   81
         Top             =   120
         Width           =   1335
      End
   End
   Begin MSDataGridLib.DataGrid dbgrid1 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   11033
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
   Begin VB.Menu ahyy1 
      Caption         =   "&Add"
   End
   Begin VB.Menu dmi22 
      Caption         =   "&Modifica"
   End
   Begin VB.Menu dfj8221 
      Caption         =   "&Borra"
   End
   Begin VB.Menu dk281 
      Caption         =   "&Reporte"
   End
   Begin VB.Menu dii33 
      Caption         =   "&Diagnostico"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu fdo33 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tconsult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsc   As New ADODB.Recordset

Dim rdiag As New ADODB.Recordset

Private Sub SQL()

    Dim xfechai As String

    Dim xfechaf As String

    On Error GoTo cmd5_err

    Dim cad As String

    If Len(xsede) = 0 Then
        xsede.SetFocus
        Exit Sub

    End If

    If Not IsDate(fechai) Then
        MsgBox "Fecha Inicio erroneo", 48, "Aviso"
        Exit Sub

    End If

    If Not IsDate(fechaf) Then
        MsgBox "Fecha Final erroneo", 48, "Aviso"
        Exit Sub

    End If

    cad = "SELECT Consulta.sede,Consulta.Consulta,consulta.Fecha,clientes.Nombre,consulta.cliente,consulta.referencia,consulta.clinica,consulta.medico,consulta.tiposeguro,consulta.empresa,consulta.tipoafilia,consulta.titularse,consulta.tipoautor,consulta.codigoauto,consulta.deducible,consulta.coaseguro,consulta.tipoconsu,consulta.medicoosi,consulta.nro,consulta.partisaldo,consulta.empresaldo FROM consulta,clientes where consulta.cliente=clientes.codigo "
    cad = cad & " and consulta.fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    cad = cad & " and consulta.fecha<='" & Format(fechaf, "YYYYMMDD") & "'"
    cad = cad & "  and consulta.sede='" & xsede & "'"

    If xcliente <> "%" Then
        cad = cad & " and consulta.cliente='" & xcliente & "'"

    End If

    cad = cad & " order by consulta.fecha"

    'MsgBox cad
    If rsc.State = 1 Then rsc.Close
    rsc.Open cad, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = rsc
    dbGrid1.columns(0).Width = 500
    dbGrid1.columns(1).Width = 800
    dbGrid1.columns(2).Width = 1000
    dbGrid1.columns(3).Width = 3900
    dbGrid1.columns(4).Width = 1000

    Exit Sub
cmd5_err:
    MsgBox "Aviso en sql " + error, 48, "Aviso"
    Exit Sub

End Sub

Sub sql1(xsede As String, xconsulta As String)

    On Error GoTo cmd8_err

    Dim cad As String

    cad = "SELECT diagnostico.diagnostico as Nro,Diagnostico.enfermedad,enfermedad.nombre,diagnostico.observa1,diagnostico.observa2,diagnostico.observa3,diagnostico.observa4 FROM diagnostico,enfermedad where diagnostico.enfermedad=enfermedad.enfermedad and diagnostico.sede='" & dsede & "' and diagnostico.consulta='" & dconsulta & "'"

    If rdiag.State = 1 Then rdiag.Close
    rdiag.Open cad, cn, adOpenStatic, adLockOptimistic
    Set dbgrid3.DataSource = rdiag
    dbgrid3.columns(0).Width = 1000
    dbgrid3.columns(1).Width = 1000
    dbgrid3.columns(2).Width = 4000
    dbgrid3.columns(3).Width = 4000
    dbgrid3.columns(4).Width = 4000
    dbgrid3.columns(5).Width = 4000
    dbgrid3.SetFocus
    Exit Sub
cmd8_err:
    MsgBox "Aviso en sql " + error, 48, "Aviso"
    Exit Sub

End Sub

Private Sub ahyy1_Click()

    Dim rsede As New ADODB.Recordset

    If Frame1.Visible = True Then Exit Sub
    If Frame3.Visible = True Then
        If Frame4.Visible = True Then
            enfermedad.SetFocus
            Exit Sub

        End If

        Frame4.Visible = True
        Frame4.Caption = "NUEVO"
        diagnostico = ""
        enfermedad = ""
        nenfermedad = ""
        observa1 = ""
        observa2 = ""
        observa3 = ""
        observa4 = ""
        enfermedad.SetFocus
        Exit Sub

    End If

    Frame1.Visible = True
    Frame1.Caption = "NUEVO"
    codigo = ""
    sede = glocal
    nsede = glocal
    sede.Enabled = False
    codigo.Enabled = False
    inicializa
    cliente.SetFocus

End Sub

Sub inicializa()
    nro = ""
    nombre = ""
    cliente = ""
    referencia = ""
    clinica = ""
    medico = ""
    tiposeguro = ""
    empresa = ""
    tipoafilia = ""
    titularse = ""
    tipoautor = ""
    codigoauto = ""
    deducible = ""
    coaseguro = ""
    fecha = Format(Now, "dd/mm/yyyy")
    tipoconsu = ""
    medicoosi = ""

    nreferencia = ""
    nclinica = ""
    nmedico = ""
    ntiposeguro = ""
    nempresa = ""
    ntipoafilia = ""
    ntitularse = ""
    ntipoautor = ""

    ntipoconsu = ""
    nmedicoosi = ""

End Sub

Private Sub buffer_DblClick()
    Command1_Click

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        If opcion1 = 1 Then
            Frame2.Visible = False
            cliente.SetFocus
            Exit Sub

        End If

        If opcion1 = 2 Then
            Frame2.Visible = False
            referencia.SetFocus
            Exit Sub

        End If

        If opcion1 = 3 Then
            Frame2.Visible = False
            clinica.SetFocus
            Exit Sub

        End If

        If opcion1 = 4 Then
            Frame2.Visible = False
            medico.SetFocus
            Exit Sub

        End If

        If opcion1 = 5 Then
            Frame2.Visible = False
            tiposeguro.SetFocus
            Exit Sub

        End If

        If opcion1 = 6 Then
            Frame2.Visible = False
            empresa.SetFocus
            Exit Sub

        End If

        If opcion1 = 7 Then
            Frame2.Visible = False
            tipoafilia.SetFocus
            Exit Sub

        End If

        If opcion1 = 8 Then
            Frame2.Visible = False
            titularse.SetFocus
            Exit Sub

        End If

        If opcion1 = 9 Then
            Frame2.Visible = False
            tipoautor.SetFocus
            Exit Sub

        End If

        If opcion1 = 10 Then
            Frame2.Visible = False
            tipoconsu.SetFocus
            Exit Sub

        End If

        If opcion1 = 11 Then
            Frame2.Visible = False
            medicoosi.SetFocus
            Exit Sub

        End If

        If opcion1 = 12 Then
            Frame2.Visible = False
            enfermedad.SetFocus
            Exit Sub

        End If

        If opcion1 = 13 Then
            Frame2.Visible = False
            xsede.SetFocus
            Exit Sub

        End If

        If opcion1 = 14 Then
            Frame2.Visible = False
            xcliente.SetFocus
            Exit Sub

        End If

    End If

    Command1_Click

End Sub

Private Sub cliente_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    fecha.SetFocus

End Sub

Private Sub cliente_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_cliente

    End If

End Sub

Sub consulta_cliente()
    Frame2.Visible = True
    Frame2.Enabled = True
    buffer = ""
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0

    opcion1 = 1
    buffer.SetFocus
    Command1_Click

End Sub

Sub consulta_Referencia()

    Dim cad       As String

    Dim rconsulta As New ADODB.Recordset

    cad = "SELECT * FROM referencia  "

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True Or rconsulta.BOF = True Then
        Exit Sub

    End If

    Frame2.Visible = True
    Frame2.Enabled = True
    buffer = ""
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Referencia"
    Combo1.ListIndex = 0
    opcion1 = 2
    buffer.SetFocus
    Command1_Click

End Sub

Sub consulta_clinica()

    Dim cad       As String

    Dim rconsulta As New ADODB.Recordset

    cad = "SELECT * FROM clinica  "

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True Or rconsulta.BOF = True Then
        Exit Sub

    End If

    Frame2.Visible = True
    Frame2.Enabled = True
    buffer = ""
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "clinica"
    Combo1.ListIndex = 0

    opcion1 = 3
    buffer.SetFocus
    Command1_Click

End Sub

Sub consulta_medico()

    Dim cad       As String

    Dim rconsulta As New ADODB.Recordset

    cad = "SELECT * FROM medico  "

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True Or rconsulta.BOF = True Then
        Exit Sub

    End If

    Frame2.Visible = True
    Frame2.Enabled = True
    buffer = ""
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Medico"
    Combo1.ListIndex = 0
    opcion1 = 4
    buffer.SetFocus
    Command1_Click

End Sub

Sub consulta_seguro()

    Dim cad       As String

    Dim rconsulta As New ADODB.Recordset

    cad = "SELECT * FROM seguro  "

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True Or rconsulta.BOF = True Then
        Exit Sub

    End If

    Frame2.Visible = True
    Frame2.Enabled = True
    buffer = ""
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Seguro"
    Combo1.ListIndex = 0
    opcion1 = 5
    buffer.SetFocus
    Command1_Click

End Sub

Sub consulta_empresa()

    Dim cad       As String

    Dim rconsulta As New ADODB.Recordset

    cad = "SELECT * FROM clientes  "

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True Or rconsulta.BOF = True Then
        Exit Sub

    End If

    Frame2.Visible = True
    Frame2.Enabled = True
    buffer = ""
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    opcion1 = 6
    buffer.SetFocus
    Command1_Click

End Sub

Sub consulta_tipoafilia()

    Dim cad       As String

    Dim rconsulta As New ADODB.Recordset

    cad = "SELECT * FROM TIPOafilia  "

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True Or rconsulta.BOF = True Then
        Exit Sub

    End If

    Frame2.Visible = True
    Frame2.Enabled = True
    buffer = ""
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "tipoafilia"
    Combo1.ListIndex = 0
    opcion1 = 7
    buffer.SetFocus
    Command1_Click

End Sub

Sub consulta_cliente1()

    Dim cad       As String

    Dim rconsulta As New ADODB.Recordset

    cad = "SELECT * FROM clientes  "

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True Or rconsulta.BOF = True Then
        Exit Sub

    End If

    Frame2.Visible = True
    Frame2.Enabled = True
    buffer = ""
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    opcion1 = 8
    buffer.SetFocus
    Command1_Click

End Sub

Sub consulta_tipoauto()

    Dim cad       As String

    Dim rconsulta As New ADODB.Recordset

    cad = "SELECT * FROM TIPOautor  "

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True Or rconsulta.BOF = True Then
        Exit Sub

    End If

    Frame2.Visible = True
    Frame2.Enabled = True
    buffer = ""
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "tipoautor"
    Combo1.ListIndex = 0
    opcion1 = 9
    buffer.SetFocus
    Command1_Click

End Sub

Sub consulta_tipoconsu()

    Dim cad       As String

    Dim rconsulta As New ADODB.Recordset

    cad = "SELECT * FROM TIPOconsu  "

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True Or rconsulta.BOF = True Then
        Exit Sub

    End If

    Frame2.Visible = True
    Frame2.Enabled = True
    buffer = ""
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "tipoconsu"
    Combo1.ListIndex = 0
    opcion1 = 10
    buffer.SetFocus
    Command1_Click

End Sub

Sub consulta_osi()

    Dim cad       As String

    Dim rconsulta As New ADODB.Recordset

    cad = "SELECT * FROM medico  "

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True Or rconsulta.BOF = True Then
        Exit Sub

    End If

    Frame2.Visible = True
    Frame2.Enabled = True
    buffer = ""
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Medico"
    Combo1.ListIndex = 0
    opcion1 = 11
    buffer.SetFocus
    Command1_Click

End Sub

Sub consulta_diagnostico()

    Dim cad       As String

    Dim rconsulta As New ADODB.Recordset

    cad = "SELECT * FROM enfermedad  "

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True Or rconsulta.BOF = True Then
        Exit Sub

    End If

    Frame2.Visible = True
    Frame2.Enabled = True
    buffer = ""
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Enfermedad"
    Combo1.ListIndex = 0
    opcion1 = 12
    buffer.SetFocus
    Command1_Click

End Sub

Sub consulta_sede()

    Dim cad       As String

    Dim rconsulta As New ADODB.Recordset

    cad = "SELECT * FROM tlocal  "

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True Or rconsulta.BOF = True Then
        Exit Sub

    End If

    Frame2.Visible = True
    Frame2.Enabled = True
    buffer = ""
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "codigo"
    Combo1.ListIndex = 0
    opcion1 = 13
    buffer.SetFocus
    Command1_Click

End Sub

Sub consulta_xcliente()

    Dim cad       As String

    Dim rconsulta As New ADODB.Recordset

    cad = "SELECT * FROM clientes  "

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True Or rconsulta.BOF = True Then
        Exit Sub

    End If

    Frame2.Visible = True
    Frame2.Enabled = True
    buffer = ""
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    opcion1 = 14
    buffer.SetFocus
    Command1_Click

End Sub

Private Sub clinica_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    medico.SetFocus

End Sub

Private Sub clinica_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_clinica

    End If

End Sub

Private Sub cmdAddEntry_Click()
    ahyy1_Click

End Sub

Private Sub cmdDelete_Click()
    dfj8221_Click

End Sub

Private Sub cmdExit_Click()
    fdo33_Click

End Sub

Private Sub cmdHelp_Click()
    dmi22_Click

End Sub

Private Sub cmdPrint_Click()
    dk281_Click

End Sub

Private Sub coaseguro_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    tipoconsu.SetFocus

End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub codigoauto_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    deducible.SetFocus

End Sub

Private Sub Command1_Click()
    ejecuta 1

End Sub

Sub ejecuta(sw As Integer)

    Dim rconsulta As New ADODB.Recordset

    Dim cad       As String

    If opcion1 = 1 Then  'clientes
        If Len(buffer) = 0 Then
            cad = "SELECT Nombre,Codigo FROM clientes  "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT Nombre,Codigo FROM clientes where  " & Combo1 & " like '" & buffer & "%'"

        End If

        If rconsulta.State = 1 Then rconsulta.Close
        rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

        If rconsulta.EOF = True And rconsulta.BOF = True Then
            buffer.SetFocus
            Exit Sub

        End If

        Set DBGrid2.DataSource = rconsulta
        DBGrid2.columns(0).Width = 5000
        DBGrid2.columns(1).Width = 1000

        If sw = 1 Then
            DBGrid2.SetFocus

        End If

    End If

    If opcion1 = 2 Then  'referencia
        If Len(buffer) = 0 Then
            cad = "SELECT Nombre,Referencia FROM Referencia  "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT Nombre,Referencia FROM Referencia where  " & Combo1 & " like '" & buffer & "%'"

        End If

        If rconsulta.State = 1 Then rconsulta.Close
        rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

        If rconsulta.EOF = True And rconsulta.BOF = True Then
            buffer.SetFocus
            Exit Sub

        End If

        Set DBGrid2.DataSource = rconsulta
        DBGrid2.columns(0).Width = 5000
        DBGrid2.columns(1).Width = 1000

        If sw = 1 Then
            DBGrid2.SetFocus

        End If

    End If

    If opcion1 = 3 Then  'clinica
        If Len(buffer) = 0 Then
            cad = "SELECT Nombre,Clinica FROM clinica  "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT Nombre,Clinica FROM clinica where  " & Combo1 & " like '" & buffer & "%'"

        End If

        If rconsulta.State = 1 Then rconsulta.Close
        rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

        If rconsulta.EOF = True And rconsulta.BOF = True Then
            buffer.SetFocus
            Exit Sub

        End If

        Set DBGrid2.DataSource = rconsulta
        DBGrid2.columns(0).Width = 5000
        DBGrid2.columns(1).Width = 1000

        If sw = 1 Then
            DBGrid2.SetFocus

        End If

    End If

    If opcion1 = 4 Then  'medico
        If Len(buffer) = 0 Then
            cad = "SELECT Nombre,Medico FROM Medico  "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT Nombre,Medico FROM Medico where  " & Combo1 & " like '" & buffer & "%'"

        End If

        If rconsulta.State = 1 Then rconsulta.Close
        rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

        If rconsulta.EOF = True And rconsulta.BOF = True Then
            buffer.SetFocus
            Exit Sub

        End If

        Set DBGrid2.DataSource = rconsulta
        DBGrid2.columns(0).Width = 5000
        DBGrid2.columns(1).Width = 1000

        If sw = 1 Then
            DBGrid2.SetFocus

        End If

    End If

    If opcion1 = 5 Then  'seguro
        If Len(buffer) = 0 Then
            cad = "SELECT Nombre,Seguro FROM seguro  "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT Nombre,Seguro FROM seguro where  " & Combo1 & " like '" & buffer & "%'"

        End If

        If rconsulta.State = 1 Then rconsulta.Close
        rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

        If rconsulta.EOF = True And rconsulta.BOF = True Then
            buffer.SetFocus
            Exit Sub

        End If

        Set DBGrid2.DataSource = rconsulta
        DBGrid2.columns(0).Width = 5000
        DBGrid2.columns(1).Width = 1000

        If sw = 1 Then
            DBGrid2.SetFocus

        End If

    End If

    If opcion1 = 6 Then  'seguro
        If Len(buffer) = 0 Then
            cad = "SELECT Nombre,codigo FROM Clientes  "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT Nombre,Codigo FROM Clientes where  " & Combo1 & " like '" & buffer & "%'"

        End If

        If rconsulta.State = 1 Then rconsulta.Close
        rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

        If rconsulta.EOF = True And rconsulta.BOF = True Then
            buffer.SetFocus
            Exit Sub

        End If

        Set DBGrid2.DataSource = rconsulta
        DBGrid2.columns(0).Width = 5000
        DBGrid2.columns(1).Width = 1000

        If sw = 1 Then
            DBGrid2.SetFocus

        End If

    End If

    If opcion1 = 7 Then  'seguro
        If Len(buffer) = 0 Then
            cad = "SELECT Nombre,Tipoafilia FROM tipoafilia  "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT Nombre,Tipoafilia FROM tipoafilia where  " & Combo1 & " like '" & buffer & "%'"

        End If

        If rconsulta.State = 1 Then rconsulta.Close
        rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

        If rconsulta.EOF = True And rconsulta.BOF = True Then
            buffer.SetFocus
            Exit Sub

        End If

        Set DBGrid2.DataSource = rconsulta
        DBGrid2.columns(0).Width = 5000
        DBGrid2.columns(1).Width = 1000

        If sw = 1 Then
            DBGrid2.SetFocus

        End If

    End If

    If opcion1 = 8 Then  'clientes
        If Len(buffer) = 0 Then
            cad = "SELECT Nombre,codigo FROM clientes  "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT Nombre,codigo FROM clientes where  " & Combo1 & " like '" & buffer & "%'"

        End If

        If rconsulta.State = 1 Then rconsulta.Close
        rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

        If rconsulta.EOF = True And rconsulta.BOF = True Then
            buffer.SetFocus
            Exit Sub

        End If

        Set DBGrid2.DataSource = rconsulta
        DBGrid2.columns(0).Width = 5000
        DBGrid2.columns(1).Width = 1000

        If sw = 1 Then
            DBGrid2.SetFocus

        End If

    End If

    If opcion1 = 9 Then  'clientes
        If Len(buffer) = 0 Then
            cad = "SELECT Nombre,tipoautor FROM tipoautor  "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT Nombre,tipoautor FROM tipoautor where  " & Combo1 & " like '" & buffer & "%'"

        End If

        If rconsulta.State = 1 Then rconsulta.Close
        rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

        If rconsulta.EOF = True And rconsulta.BOF = True Then
            buffer.SetFocus
            Exit Sub

        End If

        Set DBGrid2.DataSource = rconsulta
        DBGrid2.columns(0).Width = 5000
        DBGrid2.columns(1).Width = 1000

        If sw = 1 Then
            DBGrid2.SetFocus

        End If

    End If

    If opcion1 = 10 Then  'clientes
        If Len(buffer) = 0 Then
            cad = "SELECT Nombre,tipoconsu FROM tipoconsu  "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT Nombre,tipoconsu FROM tipoconsu where  " & Combo1 & " like '" & buffer & "%'"

        End If

        If rconsulta.State = 1 Then rconsulta.Close
        rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

        If rconsulta.EOF = True And rconsulta.BOF = True Then
            buffer.SetFocus
            Exit Sub

        End If

        Set DBGrid2.DataSource = rconsulta
        DBGrid2.columns(0).Width = 5000
        DBGrid2.columns(1).Width = 1000

        If sw = 1 Then
            DBGrid2.SetFocus

        End If

    End If

    If opcion1 = 11 Then  'clientes
        If Len(buffer) = 0 Then
            cad = "SELECT Nombre,medico FROM medico  "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT Nombre,medico FROM medico where  " & Combo1 & " like '" & buffer & "%'"

        End If

        If rconsulta.State = 1 Then rconsulta.Close
        rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

        If rconsulta.EOF = True And rconsulta.BOF = True Then
            buffer.SetFocus
            Exit Sub

        End If

        Set DBGrid2.DataSource = rconsulta
        DBGrid2.columns(0).Width = 5000
        DBGrid2.columns(1).Width = 1000

        If sw = 1 Then
            DBGrid2.SetFocus

        End If

    End If

    If opcion1 = 12 Then  'clientes
        If Len(buffer) = 0 Then
            cad = "SELECT Nombre,Enfermedad FROM enfermedad  "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT Nombre,Enfermedad FROM enfermedad where  " & Combo1 & " like '" & buffer & "%'"

        End If

        If rconsulta.State = 1 Then rconsulta.Close
        rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

        If rconsulta.EOF = True And rconsulta.BOF = True Then
            buffer.SetFocus
            Exit Sub

        End If

        Set DBGrid2.DataSource = rconsulta
        DBGrid2.columns(0).Width = 5000
        DBGrid2.columns(1).Width = 1000

        If sw = 1 Then
            DBGrid2.SetFocus

        End If

    End If

    If opcion1 = 13 Then  'clientes
        If Len(buffer) = 0 Then
            cad = "SELECT Nombre,codigo FROM tlocal  "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT Nombre,codigo FROM tlocal where  " & Combo1 & " like '" & buffer & "%'"

        End If

        If rconsulta.State = 1 Then rconsulta.Close
        rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

        If rconsulta.EOF = True And rconsulta.BOF = True Then
            buffer.SetFocus
            Exit Sub

        End If

        Set DBGrid2.DataSource = rconsulta
        DBGrid2.columns(0).Width = 5000
        DBGrid2.columns(1).Width = 1000

        If sw = 1 Then
            DBGrid2.SetFocus

        End If

    End If

    If opcion1 = 14 Then  'clientes
        If Len(buffer) = 0 Then
            cad = "SELECT Nombre,codigo FROM clientes  "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT Nombre,codigo FROM clientes where  " & Combo1 & " like '" & buffer & "%'"

        End If

        If rconsulta.State = 1 Then rconsulta.Close
        rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

        If rconsulta.EOF = True And rconsulta.BOF = True Then
            buffer.SetFocus
            Exit Sub

        End If

        Set DBGrid2.DataSource = rconsulta
        DBGrid2.columns(0).Width = 5000
        DBGrid2.columns(1).Width = 1000

        If sw = 1 Then
            DBGrid2.SetFocus

        End If

    End If

End Sub

Private Sub Command2_Click()
    SQL
    dbGrid1.SetFocus

End Sub

Private Sub Command3_Click()

    Dim found    As Integer

    Dim rs1      As New ADODB.Recordset

    Dim rsexiste As New ADODB.Recordset

    Dim cad      As String

    Dim sdx      As Double

    Dim hora     As String

    On Error GoTo cmd2_err

    found = valida()

    If found = 0 Then
        MsgBox "Campos invalidos ", 48, "Aviso"
        Exit Sub

    End If

    If Frame1.Caption = "NUEVO" Then
        hora = Format(Now, "hh:mm:ss")

        'conseguimos el numero
        If rs1.State = 1 Then rs1.Close
        rs1.Open "SELECT numeroc FROM parame where codigo='01' ", cn, adOpenKeyset, adLockOptimistic

        If rs1.RecordCount = 0 Then  'si existe
            MsgBox "No hay Sedes  ", 48, "Aviso"
            cliente.SetFocus
            Exit Sub

        End If

        sdx = Val("" & rs1.Fields("numeroc").Value) + 1
siguen:
        codigo = "" & sdx

        If rsexiste.State = 1 Then rsexiste.Close
        rsexiste.Open "SELECT consulta FROM consulta where consulta='" & Trim(codigo) & "' and sede='" & sede & "'", cn, adOpenKeyset, adLockOptimistic

        If rsexiste.RecordCount > 0 Then  'si existe
            sdx = sdx + 1
            GoTo siguen
            Exit Sub

        End If

        cad = "update PARAME set numeroc='" & codigo & "' where codigo='01'"
        cn.Execute (cad)
        cad = "INSERT INTO consulta VALUES('" & Trim(codigo) & "','" & Trim(cliente) & "','" & Trim(referencia) & "','" & Trim(clinica) & "','" & Trim(medico) & "','" & Trim(tiposeguro) & "','" & Trim(titularse) & "','" & Trim(tipoautor) & "','" & Trim(codigoauto) & "'," & Val(deducible) & "," & Val(coaseguro) & ",'" & Trim(fecha) & "','" & Trim(hora) & "','" & Trim(tipoconsu) & "','" & Trim(medicoosi) & "','" & Trim(sede) & "','" & Trim(empresa) & "','" & Trim(tipoafilia) & "','" & Trim(nro) & "'," & Val(partisaldo) & "," & Val(empresaldo) & ")"
        cn.Execute (cad)
        SQL
        dbGrid1.SetFocus
        fdo33_Click

    End If

    If Frame1.Caption = "MODIFICA" Then
        hora = Format(Now, "hh:mm:ss")
        cad = "UPDATE consulta SET cliente = '" & Trim(cliente) & "',referencia = '" & Trim(referencia) & "',clinica = '" & Trim(clinica) & "',medico = '" & Trim(medico) & "',tiposeguro = '" & Trim(tiposeguro) & "',titularse = '" & Trim(titularse) & "',tipoautor = '" & Trim(tipoautor) & "',codigoauto = '" & Trim(codigoauto) & "',deducible = " & Val(deducible) & ",coaseguro = " & Val(coaseguro) & ",fecha = '" & Trim(fecha) & "',hora = '" & Trim(hora) & "',tipoconsu = '" & Trim(tipoconsu) & "',medicoosi = '" & Trim(medicoosi) & "',sede = '" & Trim(sede) & "',empresa = '" & Trim(empresa) & "',tipoafilia = '" & Trim(tipoafilia) & "' WHERE consulta = '" & Trim(codigo) & "' and sede='" & sede & "'"
        cn.Execute (cad)
        SQL
        dbGrid1.SetFocus
        fdo33_Click

    End If

    Exit Sub
cmd2_err:
    MsgBox "Aviso en command3 " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command4_Click()
    fdo33_Click

End Sub

Function valida()

    If existe_codigo("" & cliente) = 0 Then
        MsgBox "No existe Cliente"
        cliente.SetFocus
        Exit Function

    End If

    valida = 1

End Function

Private Sub Command5_Click()

    Dim sdx      As Double

    Dim cad      As String

    Dim rs1      As New ADODB.Recordset

    Dim rsexiste As New ADODB.Recordset

    If Frame4.Caption = "NUEVO" Then

        'conseguimos el numero
        If rs1.State = 1 Then rs1.Close
        rs1.Open "SELECT Codigo FROM parame where codigo='01' ", cn, adOpenKeyset, adLockOptimistic

        If rs1.RecordCount = 0 Then  'si existe
            MsgBox "No hay Sedes  ", 48, "Aviso"
            cliente.SetFocus
            Exit Sub

        End If

        sdx = Val("" & rs1.Fields("numerod").Value) + 1
siguen1:
        diagnostico = "" & sdx

        If rsexiste.State = 1 Then rsexiste.Close
        rsexiste.Open "SELECT diagnostico FROM diagnostico where consulta='" & Trim(dconsulta) & "' and sede='" & dsede & "' and diagnostico='" & diagnostico & "'", cn, adOpenKeyset, adLockOptimistic

        If rsexiste.RecordCount > 0 Then  'si existe
            sdx = sdx + 1
            GoTo siguen1
            Exit Sub

        End If

        cad = "update parame set numerod='" & diagnostico & "' where codigo='01'"
        cn.Execute (cad)
        cad = "INSERT INTO diagnostico VALUES('" & Trim(diagnostico) & "','" & Trim(dsede) & "','" & Trim(dconsulta) & "','" & Trim(enfermedad) & "','" & Trim(observa1) & "','" & Trim(observa2) & "','" & Trim(observa3) & "','" & Trim(observa4) & "')"
        cn.Execute (cad)
        sql1 dsede, dconsulta
        Frame4.Visible = False
        dbgrid3.SetFocus
        Exit Sub

    End If

    If Frame4.Caption = "MODIFICA" Then
        cad = "UPDATE diagnostico SET enfermedad = '" & Trim(enfermedad) & "',observa1 = '" & Trim(observa1) & "',observa2 = '" & Trim(observa2) & "',observa3 = '" & Trim(observa3) & "',observa4 = '" & Trim(observa4) & "' WHERE consulta = '" & Trim(dconsulta) & "' and sede='" & dsede & "' and diagnostico='" & diagnostico & "'"
        cn.Execute (cad)
        sql1 dsede, dconsulta
        Frame4.Visible = False
        dbgrid3.SetFocus
        Exit Sub

    End If

End Sub

Private Sub Command6_Click()
    fdo33_Click

End Sub

Private Sub dbgrid2_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = 1 Then
            cliente = Trim(DBGrid2.columns(1))
            nombre = Trim(DBGrid2.columns(0))
            Frame2.Visible = False
            Frame2.Enabled = False
            referencia.SetFocus
      
            Exit Sub

        End If

        If opcion1 = 2 Then
            referencia = Trim(DBGrid2.columns(1))
            nreferencia = Trim(DBGrid2.columns(0))
            Frame2.Visible = False
            Frame2.Enabled = False
            clinica.SetFocus
            Exit Sub

        End If

        If opcion1 = 3 Then
            clinica = Trim(DBGrid2.columns(1))
            nclinica = Trim(DBGrid2.columns(0))
            Frame2.Visible = False
            Frame2.Enabled = False
            medico.SetFocus
            Exit Sub

        End If

        If opcion1 = 4 Then
            medico = Trim(DBGrid2.columns(1))
            nmedico = Trim(DBGrid2.columns(0))
            Frame2.Visible = False
            Frame2.Enabled = False
            tiposeguro.SetFocus
            Exit Sub

        End If

        If opcion1 = 5 Then
            tiposeguro = Trim(DBGrid2.columns(1))
            ntiposeguro = Trim(DBGrid2.columns(0))
            Frame2.Visible = False
            Frame2.Enabled = False
            empresa.SetFocus
            Exit Sub

        End If

        If opcion1 = 6 Then
            empresa = Trim(DBGrid2.columns(1))
            nempresa = Trim(DBGrid2.columns(0))
            Frame2.Visible = False
            Frame2.Enabled = False
            tipoafilia.SetFocus
            Exit Sub

        End If

        If opcion1 = 7 Then
            tipoafilia = Trim(DBGrid2.columns(1))
            ntipoafilia = Trim(DBGrid2.columns(0))
            Frame2.Visible = False
            Frame2.Enabled = False
            titularse.SetFocus
            Exit Sub

        End If

        If opcion1 = 8 Then
            titularse = Trim(DBGrid2.columns(1))
            ntitularse = Trim(DBGrid2.columns(0))
            Frame2.Visible = False
            Frame2.Enabled = False
            tipoautor.SetFocus
            Exit Sub

        End If

        If opcion1 = 9 Then
            tipoautor = Trim(DBGrid2.columns(1))
            ntipoautor = Trim(DBGrid2.columns(0))
            Frame2.Visible = False
            Frame2.Enabled = False
            codigoauto.SetFocus
            Exit Sub

        End If

        If opcion1 = 10 Then
            tipoconsu = Trim(DBGrid2.columns(1))
            ntipoconsu = Trim(DBGrid2.columns(0))
            Frame2.Visible = False
            Frame2.Enabled = False
            medicoosi.SetFocus
            Exit Sub

        End If

        If opcion1 = 11 Then
            medicoosi = Trim(DBGrid2.columns(1))
            nmedicoosi = Trim(DBGrid2.columns(0))
            Frame2.Visible = False
            Frame2.Enabled = False
            medicoosi.SetFocus
            Exit Sub

        End If

        If opcion1 = 12 Then
            enfermedad = Trim(DBGrid2.columns(1))
            nenfermedad = Trim(DBGrid2.columns(0))
            Frame2.Visible = False
            Frame2.Enabled = False
            observa1.SetFocus
            Exit Sub

        End If

        If opcion1 = 13 Then
            xsede = Trim(DBGrid2.columns(1))
            Frame2.Visible = False
            Frame2.Enabled = False
            xsede.SetFocus
            Exit Sub

        End If

        If opcion1 = 14 Then
            xcliente = Trim(DBGrid2.columns(1))
            Frame2.Visible = False
            Frame2.Enabled = False
            xcliente.SetFocus
            Exit Sub

        End If

    End If

End Sub

Private Sub dbgrid2_KeyPress(KeyAscii As Integer)

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

Private Sub deducible_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    coaseguro.SetFocus

End Sub

Private Sub dfj8221_Click()

    Dim buf As String

    On Error GoTo cmd4_err

    If Frame1.Visible = True Then Exit Sub
    If Frame3.Visible = True Then
        If Frame4.Visible = True Then
            enfermedad.SetFocus
            Exit Sub

        End If

        buf = Trim(dbGrid1.columns(1))

        If MsgBox("Desea Borrar " + dbgrid3.columns(0), 1, "Aviso") = 1 Then
            cn.Execute ("DELETE   FROM diagnostico WHERE consulta ='" & Trim(dconsulta) & "' and sede='" & Trim(dsede) & "' and diagnostico='" & Trim(dbgrid3.columns(0)) & "'")
            rdiag.Requery
            sql1 dsede, dconsulta

        End If

        dbgrid3.SetFocus
        Exit Sub

    End If

    buf = Trim(dbGrid1.columns(1))

    If MsgBox("Desea Borrar " + dbGrid1.columns(1), 1, "Aviso") = 1 Then
        cn.Execute ("DELETE   FROM consulta WHERE consulta ='" & Trim(dbGrid1.columns(1)) & "' and sede='" & Trim(dbGrid1.columns(0)) & "'")
        rsc.Requery
        SQL
        dbGrid1.SetFocus

    End If

    dbGrid1.SetFocus
    Exit Sub
cmd4_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    dbGrid1.SetFocus
    Exit Sub

End Sub

Private Sub dii33_Click()

    On Error GoTo cmd10_err

    If Frame1.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    dsede = Trim(dbGrid1.columns(0))
    dconsulta = Trim(dbGrid1.columns(1))
    Frame3.Visible = True
    sql1 dsede, dconsulta

    Exit Sub
cmd10_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub dk281_Click()
    'If Frame1.Visible = True Then Exit Sub
    'If rt.State = 1 Then rt.Close
    'rt.Open "SELECT * FROM consulta ", cn, adOpenKeyset, adLockOptimistic
    'Set trepcli1.DataSource = rt
    'trepcli1.Show 1

End Sub

Private Sub dmi22_Click()

    Dim rs1 As New ADODB.Recordset

    On Error GoTo cmd3_err

    If Frame1.Visible = True Then Exit Sub
    If Frame3.Visible = True Then
        If Frame4.Visible = True Then
            enfermedad.SetFocus
            Exit Sub

        End If

        diagnostico = Trim(dbGrid1.columns(0))
        enfermedad = Trim(dbGrid1.columns(1))
        nenfermedad = Trim(dbGrid1.columns(2))
        observa1 = Trim(dbGrid1.columns(3))
        observa2 = Trim(dbGrid1.columns(4))
        observa3 = Trim(dbGrid1.columns(5))
        observa4 = Trim(dbGrid1.columns(6))
        Frame4.Visible = True
        Frame4.Caption = "MODIFICA"
        diagnostico.Enabled = False
        enfermedad.SetFocus
        Exit Sub

    End If

    inicializa
    cliente = Trim(dbGrid1.columns(4))
    sede = Trim(dbGrid1.columns(0))
    codigo = Trim(dbGrid1.columns(1))
    fecha = Trim(dbGrid1.columns(2))
    referencia = Trim(dbGrid1.columns(5))
    clinica = Trim(dbGrid1.columns(6))
    medico = Trim(dbGrid1.columns(7))
    tiposeguro = Trim(dbGrid1.columns(8))
    empresa = Trim(dbGrid1.columns(9))
    tipoafilia = Trim(dbGrid1.columns(10))
    titularse = Trim(dbGrid1.columns(11))
    tipoautor = Trim(dbGrid1.columns(12))
    codigoauto = Trim(dbGrid1.columns(13))
    deducible = Trim(dbGrid1.columns(14))
    coaseguro = Trim(dbGrid1.columns(15))
    tipoconsu = Trim(dbGrid1.columns(16))
    medicoosi = Trim(dbGrid1.columns(17))
    nro = Trim(dbGrid1.columns(18))
    partisaldo = Trim(dbGrid1.columns(19))
    empresaldo = Trim(dbGrid1.columns(20))

    'busquedas
    If rs1.State = 1 Then rs1.Close
    rs1.Open "SELECT Nombre FROM clientes where codigo='" & Trim(cliente) & "'", cn, adOpenKeyset, adLockOptimistic

    If rs1.RecordCount > 0 Then  'si existe
        nombre = "" & rs1.Fields("nombre").Value

    End If

    If rs1.State = 1 Then rs1.Close
    rs1.Open "SELECT Nombre FROM referencia where referencia='" & Trim(referencia) & "'", cn, adOpenKeyset, adLockOptimistic

    If rs1.RecordCount > 0 Then  'si existe
        nreferencia = "" & rs1.Fields("nombre").Value

    End If

    If rs1.State = 1 Then rs1.Close
    rs1.Open "SELECT nombre FROM clinica where clinica='" & Trim(clinica) & "'", cn, adOpenKeyset, adLockOptimistic

    If rs1.RecordCount > 0 Then  'si existe
        nclinica = "" & rs1.Fields("nombre").Value

    End If

    If rs1.State = 1 Then rs1.Close
    rs1.Open "SELECT nombre FROM medico where medico='" & Trim(medico) & "'", cn, adOpenKeyset, adLockOptimistic

    If rs1.RecordCount > 0 Then  'si existe
        nmedico = "" & rs1.Fields("nombre").Value

    End If

    If rs1.State = 1 Then rs1.Close
    rs1.Open "SELECT Nombre FROM seguro where seguro='" & Trim(tiposeguro) & "'", cn, adOpenKeyset, adLockOptimistic

    If rs1.RecordCount > 0 Then  'si existe
        ntiposeguro = "" & rs1.Fields("nombre").Value

    End If

    If rs1.State = 1 Then rs1.Close
    rs1.Open "SELECT Nombre FROM clientes where codigo='" & Trim(empresa) & "'", cn, adOpenKeyset, adLockOptimistic

    If rs1.RecordCount > 0 Then  'si existe
        nempresa = "" & rs1.Fields("nombre").Value

    End If

    If rs1.State = 1 Then rs1.Close
    rs1.Open "SELECT Nombre FROM tipoafilia where tipoafilia='" & Trim(tipoafilia) & "'", cn, adOpenKeyset, adLockOptimistic

    If rs1.RecordCount > 0 Then  'si existe
        ntipoafilia = "" & rs1.Fields("nombre").Value

    End If

    If rs1.State = 1 Then rs1.Close
    rs1.Open "SELECT Nombre FROM clientes where codigo='" & Trim(titularse) & "'", cn, adOpenKeyset, adLockOptimistic

    If rs1.RecordCount > 0 Then  'si existe
        ntitularse = "" & rs1.Fields("nombre").Value

    End If

    If rs1.State = 1 Then rs1.Close
    rs1.Open "SELECT Nombre FROM tipoautor where tipoautor='" & Trim(tipoautor) & "'", cn, adOpenKeyset, adLockOptimistic

    If rs1.RecordCount > 0 Then  'si existe
        ntipoautor = "" & rs1.Fields("nombre").Value

    End If

    If rs1.State = 1 Then rs1.Close
    rs1.Open "SELECT Nombre FROM tipoconsu where tipoconsu='" & Trim(tipoconsu) & "'", cn, adOpenKeyset, adLockOptimistic

    If rs1.RecordCount > 0 Then  'si existe
        ntipoconsu = "" & rs1.Fields("nombre").Value

    End If

    If rs1.State = 1 Then rs1.Close
    rs1.Open "SELECT Nombre FROM medico where medico='" & Trim(medicoosi) & "'", cn, adOpenKeyset, adLockOptimistic

    If rs1.RecordCount > 0 Then  'si existe
        nmedicoosi = "" & rs1.Fields("nombre").Value

    End If

    Frame1.Visible = True
    Frame1.Caption = "MODIFICA"
    codigo.Enabled = False
    cliente.SetFocus
    Exit Sub
cmd3_err:
    MsgBox "Seleccione un dato " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub empresa_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    tipoafilia.SetFocus

End Sub

Private Sub empresa_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_empresa

    End If

End Sub

Private Sub enfermedad_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    observa1.SetFocus

End Sub

Private Sub enfermedad_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_diagnostico

    End If

End Sub

Private Sub fdo33_Click()

    If Frame2.Visible = True Then
        buffer_KeyPress 27
        Exit Sub

    End If

    If Frame3.Visible = True Then
        If Frame4.Visible = False Then
            Frame3.Visible = False
            dbGrid1.SetFocus
            Exit Sub

        End If

        If Frame4.Caption = "NUEVO" Then
            Frame4.Visible = False
            dbgrid3.SetFocus

        End If

        If Frame4.Caption = "MODIFICA" Then
            Frame4.Visible = False
            dbgrid3.SetFocus

        End If
   
        Exit Sub

    End If

    If Frame1.Visible = True Then
        If Frame1.Caption = "NUEVO" Then
            Frame1.Visible = False
            dbGrid1.SetFocus

        End If

        If Frame1.Caption = "MODIFICA" Then
            Frame1.Visible = False
            dbGrid1.SetFocus

        End If

        Exit Sub

    End If

    tconsult.Hide
    Unload tconsult

End Sub

Private Sub fecha_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    referencia.SetFocus

End Sub

Private Sub Form_Load()
    Frame1.Top = 10: Frame1.Left = 10
    Frame2.Top = 10: Frame2.Left = 10
    Frame3.Top = 10: Frame3.Left = 10

    fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    fechaf = Format(Now, "dd/mm/yyyy")
    xsede = glocal
    SQL

End Sub

Private Sub nombre_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub Label4_Click()

    'tcli.Show 1
End Sub

Private Sub medico_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    tiposeguro.SetFocus

End Sub

Private Sub medico_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_medico

    End If

End Sub

Private Sub medicoosi_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub medicoosi_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_osi

    End If

End Sub

Private Sub observa1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    observa2.SetFocus

End Sub

Private Sub observa2_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    observa3.SetFocus

End Sub

Private Sub observa3_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    observa4.SetFocus

End Sub

Private Sub observa4_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub referencia_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    clinica.SetFocus

End Sub

Private Sub referencia_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_Referencia

    End If

End Sub

Private Sub tipoafilia_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    titularse.SetFocus

End Sub

Private Sub tipoafilia_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_tipoafilia

    End If

End Sub

Private Sub tipoautor_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    codigoauto.SetFocus

End Sub

Private Sub tipoautor_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_tipoauto

    End If

End Sub

Private Sub tipoconsu_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    medicoosi.SetFocus

End Sub

Private Sub tipoconsu_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_tipoconsu

    End If

End Sub

Private Sub tiposeguro_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    empresa.SetFocus

End Sub

Private Sub tiposeguro_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_seguro

    End If

End Sub

Private Sub titularse_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    tipoautor.SetFocus

End Sub

Private Sub titularse_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_cliente1

    End If

End Sub

Private Sub xcliente_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    SQL
    dbGrid1.SetFocus

End Sub

Private Sub xcliente_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_xcliente

    End If

End Sub

Private Sub xsede_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    xcliente.SetFocus

End Sub

Private Sub xsede_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_sede

    End If

End Sub

Function existe_codigo(buf As String)

    Dim rs1 As New ADODB.Recordset

    existe_codigo = 1

    If rs1.State = 1 Then rs1.Close
    rs1.Open "SELECT codigo FROM clientes where codigo='" & buf & "'", cn, adOpenDynamic, adLockReadOnly

    If rs1.EOF Then
        existe_codigo = 0

    End If

    rs1.Close
    Set rs1 = Nothing

End Function

