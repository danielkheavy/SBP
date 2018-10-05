VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form texplTRA 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Letras"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   16065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   16065
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Pagos"
      Height          =   7935
      Left            =   3120
      TabIndex        =   53
      Top             =   0
      Visible         =   0   'False
      Width           =   14655
      Begin MSDataGridLib.DataGrid dbgrid3 
         Height          =   5655
         Left            =   120
         TabIndex        =   54
         Top             =   1440
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   9975
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
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
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cerrar"
         Height          =   495
         Left            =   13320
         TabIndex        =   72
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BorraRecibo"
         Enabled         =   0   'False
         Height          =   495
         Left            =   13320
         TabIndex        =   71
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label hsaldo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   7440
         TabIndex        =   70
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saldo"
         Height          =   375
         Left            =   6480
         TabIndex        =   69
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero"
         Height          =   375
         Left            =   3120
         TabIndex        =   68
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Moneda"
         Height          =   375
         Left            =   3120
         TabIndex        =   67
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         Height          =   375
         Left            =   3120
         TabIndex        =   66
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
         Height          =   375
         Left            =   600
         TabIndex        =   65
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         Height          =   375
         Left            =   600
         TabIndex        =   64
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie"
         Height          =   375
         Left            =   600
         TabIndex        =   63
         Top             =   960
         Width           =   975
      End
      Begin VB.Label tsdx1 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6840
         TabIndex        =   62
         Top             =   7200
         Width           =   1695
      End
      Begin VB.Label tsdx 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4800
         TabIndex        =   61
         Top             =   7200
         Width           =   1695
      End
      Begin VB.Label htotal 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   60
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label hmoneda 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   59
         Top             =   600
         Width           =   735
      End
      Begin VB.Label hnumero 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   58
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label hserie 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1560
         TabIndex        =   57
         Top             =   960
         Width           =   975
      End
      Begin VB.Label htipo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1560
         TabIndex        =   56
         Top             =   600
         Width           =   975
      End
      Begin VB.Label hlocal 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1560
         TabIndex        =   55
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Impresion"
      Height          =   3855
      Left            =   3480
      TabIndex        =   26
      Top             =   1800
      Visible         =   0   'False
      Width           =   5535
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
         Left            =   3720
         MaskColor       =   &H00FFFFFF&
         Picture         =   "texplcxc.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.ComboBox producto 
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
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Productos"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFC0&
      Height          =   1275
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   16005
      TabIndex        =   0
      Top             =   0
      Width           =   16065
      Begin VB.ComboBox local1 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox servicios 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   11640
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox condicion 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   9240
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox codigo 
         Height          =   375
         Left            =   9240
         MaxLength       =   11
         TabIndex        =   20
         Text            =   "%"
         Top             =   0
         Width           =   1815
      End
      Begin VB.TextBox nombre 
         Height          =   375
         Left            =   9240
         MaxLength       =   11
         TabIndex        =   19
         Text            =   "%"
         Top             =   360
         Width           =   1815
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
         Height          =   735
         Left            =   13320
         MaskColor       =   &H00FFFFFF&
         Picture         =   "texplcxc.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1335
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
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox caja 
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
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   0
         Width           =   1335
      End
      Begin VB.ComboBox cajero 
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
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox fechaf 
         Height          =   375
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox fechai 
         Height          =   375
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   5
         Top             =   0
         Width           =   1815
      End
      Begin VB.ComboBox tipo 
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
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   0
         Width           =   1815
      End
      Begin VB.ComboBox vendedor 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox ordenado 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cta"
         Height          =   375
         Left            =   11040
         TabIndex        =   38
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Condicion Saldo"
         Height          =   375
         Left            =   7920
         TabIndex        =   25
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
         Height          =   375
         Left            =   5280
         TabIndex        =   23
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         Height          =   375
         Left            =   7920
         TabIndex        =   22
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         Height          =   375
         Left            =   7920
         TabIndex        =   21
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cajero"
         Height          =   375
         Left            =   2880
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja"
         Height          =   375
         Left            =   2880
         TabIndex        =   16
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Turno"
         Height          =   375
         Left            =   2880
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   375
         Left            =   0
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoDocto"
         Height          =   375
         Left            =   5280
         TabIndex        =   12
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vendedor"
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ordenado Por"
         Height          =   375
         Left            =   5280
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
   End
   Begin MSDataGridLib.DataGrid dbgrid2 
      Height          =   6015
      Left            =   0
      TabIndex        =   51
      Top             =   1320
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   10610
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
      ColumnCount     =   26
      BeginProperty Column00 
         DataField       =   "X"
         Caption         =   "X"
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
      BeginProperty Column05 
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
      BeginProperty Column06 
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
      BeginProperty Column07 
         DataField       =   "Tipoclie"
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
         DataField       =   "Cuota"
         Caption         =   "Cuota"
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
      BeginProperty Column12 
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
      BeginProperty Column13 
         DataField       =   "Interes"
         Caption         =   "Interes"
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
      BeginProperty Column14 
         DataField       =   "Abono"
         Caption         =   "Abono"
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
      BeginProperty Column15 
         DataField       =   "Saldo"
         Caption         =   "Saldo"
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
      BeginProperty Column16 
         DataField       =   "Fechav"
         Caption         =   "Fechav"
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
         DataField       =   "Grupo"
         Caption         =   "Grupo"
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
      BeginProperty Column18 
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
      BeginProperty Column19 
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
      BeginProperty Column20 
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
      BeginProperty Column21 
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
      BeginProperty Column22 
         DataField       =   "anticipo"
         Caption         =   "Anticipo"
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
      BeginProperty Column23 
         DataField       =   "Zona"
         Caption         =   "Zona"
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
         DataField       =   "fpago"
         Caption         =   "Acu"
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
      BeginProperty Column25 
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
            ColumnWidth     =   195.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   225.071
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1094.74
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
            ColumnWidth     =   464.882
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   390.047
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   195.024
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column22 
         EndProperty
         BeginProperty Column23 
         EndProperty
         BeginProperty Column24 
         EndProperty
         BeginProperty Column25 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Depo.bancario"
      Height          =   375
      Left            =   8640
      TabIndex        =   50
      Top             =   8880
      Width           =   1215
   End
   Begin VB.Label totalh 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   9840
      TabIndex        =   49
      Top             =   8880
      Width           =   1695
   End
   Begin VB.Label saldoh 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   13200
      TabIndex        =   48
      Top             =   8880
      Width           =   1575
   End
   Begin VB.Label abonoh 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11520
      TabIndex        =   47
      Top             =   8880
      Width           =   1695
   End
   Begin VB.Label abonoo 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11520
      TabIndex        =   46
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Label saldoo 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   13200
      TabIndex        =   45
      Top             =   8520
      Width           =   1575
   End
   Begin VB.Label totalo 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   9840
      TabIndex        =   44
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OrdenTrabajo"
      Height          =   375
      Left            =   8640
      TabIndex        =   43
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Label abonoa 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11520
      TabIndex        =   42
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Label saldoa 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   13200
      TabIndex        =   41
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label totala 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   9840
      TabIndex        =   40
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Adelantos"
      Height          =   375
      Left            =   8640
      TabIndex        =   39
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Label Label30 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Creditos"
      Height          =   375
      Left            =   8640
      TabIndex        =   36
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Label abonoc 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11520
      TabIndex        =   35
      Top             =   7800
      Width           =   1695
   End
   Begin VB.Label Label28 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo"
      Height          =   375
      Left            =   13200
      TabIndex        =   34
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Label Label27 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Abonos"
      Height          =   375
      Left            =   11520
      TabIndex        =   33
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label Label26 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      Height          =   375
      Left            =   9840
      TabIndex        =   32
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label saldoc 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   13200
      TabIndex        =   31
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label totalc 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   9840
      TabIndex        =   30
      Top             =   7800
      Width           =   1695
   End
   Begin VB.Label acu 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   16680
      TabIndex        =   1
      Top             =   1920
      Width           =   105
   End
   Begin VB.Menu ldo232 
      Caption         =   "&Nuevo"
   End
   Begin VB.Menu mofdi782 
      Caption         =   "&Modifica"
   End
   Begin VB.Menu cno8923 
      Caption         =   "&Zoom"
   End
   Begin VB.Menu dj333 
      Caption         =   "&Borra"
   End
   Begin VB.Menu KU7zom8 
      Caption         =   "&VerPagos"
   End
   Begin VB.Menu dj7823 
      Caption         =   "&Pagar"
   End
   Begin VB.Menu ncu773 
      Caption         =   "&Canje"
      Visible         =   0   'False
   End
   Begin VB.Menu dk8rep3 
      Caption         =   "&Reporte"
      Begin VB.Menu lk993 
         Caption         =   "&1.Normal"
      End
      Begin VB.Menu dl993 
         Caption         =   "&2.Excell"
      End
   End
   Begin VB.Menu dlo23211 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "texplTRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xcuentaco As String
Dim rxconsulta As New adodb.Recordset
Private Sub cmdCancelar_Click()

End Sub

Private Sub cmdDelete_Click()

End Sub

Private Sub cmdExit_Click()

End Sub

Private Sub cmdGrabar_Click()
End Sub

Private Sub cmdPrint_Click()

End Sub

Private Sub cmdSort_Click()
End Sub

Private Sub cno8923_Click()
On Error GoTo cmd3452_err
If Frame1.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub


tincxc.bandera = "Ver"
tincxc.acu = acu
tincxc.Local1.Enabled = False
tincxc.tipo.Enabled = False
tincxc.serie.Enabled = False
tincxc.numero.Enabled = False
tincxc.cuota.Enabled = False

tincxc.anticipo = "" & rxconsulta.Fields("anticipo")
tincxc.Local1 = "" & rxconsulta.Fields("local")

tincxc.tipo = "" & rxconsulta.Fields("tipo")
tincxc.usuario = "" & rxconsulta.Fields("usuario")

tincxc.caja = "" & rxconsulta.Fields("caja")
tincxc.turno = "" & rxconsulta.Fields("turno")
tincxc.grupo = "" & rxconsulta.Fields("grupo")



tincxc.serie = "" & rxconsulta.Fields("serie")
tincxc.numero = "" & rxconsulta.Fields("numero")
tincxc.cuota = "" & rxconsulta.Fields("cuota")
tincxc.tipoclie = "" & rxconsulta.Fields("tipoclie")
tincxc.codigo = "" & rxconsulta.Fields("codigo")
tincxc.nombre = "" & rxconsulta.Fields("nombre")
tincxc.zona = "" & rxconsulta.Fields("zona")
tincxc.vendedor = "" & rxconsulta.Fields("vendedor")
tincxc.moneda = "" & rxconsulta.Fields("moneda")
tincxc.total = "" & rxconsulta.Fields("total")
tincxc.interes = "" & rxconsulta.Fields("interes")
tincxc.abono = "" & rxconsulta.Fields("abono")
tincxc.saldo = "" & rxconsulta.Fields("saldo")
tincxc.fecha = "" & rxconsulta.Fields("fecha")
tincxc.fechav = "" & rxconsulta.Fields("fechav")

tincxc.jui12.Enabled = False
tincxc.cmdSave.Enabled = False
tincxc.Show 1
Exit Sub
cmd3452_err:
MsgBox "Seleccione un registro ", 48, "Aviso"
Exit Sub

End Sub

Private Sub Command2_Click()
sql_letras
End Sub

Private Sub Command1_Click()
dlo23211_Click
End Sub

Private Sub Command3_Click()
Dim found As Integer
Dim buf As String
Frame2.Visible = False
contpag = 0
contlin = 0
    
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
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Private Sub dbgrid2_DblClick()
On Error GoTo cmd43_err
If Trim("" & rxconsulta.Fields("x")) <> "S" Then
   rxconsulta.Fields("X") = "S"
   rxconsulta.Update
   Exit Sub
End If
If "" & rxconsulta.Fields("x") = "S" Then
   rxconsulta.Fields("X") = ""
   rxconsulta.Update
   Exit Sub
End If

Exit Sub
cmd43_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub
End Sub

Private Sub dj333_Click()
Dim found As Integer
On Error GoTo cmd566_err
If Frame1.Visible = True Then Exit Sub
found = valida_factura()
If found = 1 Then
   MsgBox "No se puede borra viene de documento ", 48, "Aviso"
   Exit Sub
End If
found = valida_recibo()
If found = 1 Then
   MsgBox "No se puede borra Ya tiene recibo ", 48, "Aviso"
   Exit Sub
End If
If MsgBox("Desea borra el documento " + rxconsulta.Fields("numero"), 1, "Aviso") <> 1 Then Exit Sub
cn.Execute ("delete from " & xcuentaco & " where local='" & rxconsulta.Fields("local") & "' and tipo='" & rxconsulta.Fields("tipo") & "' and serie='" & rxconsulta.Fields("serie") & "' and numero='" & rxconsulta.Fields("numero") & "'")
sql_letras
Exit Sub
cmd566_err:
MsgBox "Seleccione un documento ", 48, "Aviso"
Exit Sub
End Sub

Private Sub dj7823_Click()
Dim mytablex As Table
Dim found As Integer
Dim sw As Integer
On Error GoTo cmd234_err
'If local1 = "%" Then
'   MsgBox "Seleccione local ", 48, "Aviso"
'   Exit Sub
'End If
If Frame1.Visible = True Then Exit Sub

If Local1 = "%" Then
   MsgBox "Seleccione un Local ", 48, "Aviso"
   Exit Sub
End If
'preparando los recibo
   gofpago = "fpagov"
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
'ahora copiando en el temporal
sw = 0
Set mytablex = mydbxglo.OpenTable("_r" & gusuario)
If acu = "V" Then
   rxconsulta.MoveFirst
   Do
   If rxconsulta.EOF Then Exit Do
   
   If "" & rxconsulta.Fields("x") = "S" Then
   If Val("" & rxconsulta.Fields("saldo")) > 0 Then
   trecaja.tipoclie = Trim("" & rxconsulta.Fields("tipoclie"))
   trecaja.codigo = Trim("" & rxconsulta.Fields("codigo"))
   trecaja.nombre = Trim("" & rxconsulta.Fields("nombre"))
   trecaja.moneda = Trim("" & rxconsulta.Fields("moneda"))
   trecaja.Local1 = Trim("" & rxconsulta.Fields("local"))
   
   sw = 1
   mytablex.AddNew
   mytablex.Fields("codigo") = "" & rxconsulta.Fields("codigo")
   mytablex.Fields("fecha") = Format(Now, "dd/mm/yyyy")
   mytablex.Fields("hora") = Format(Now, "HH:MM:SS")
   mytablex.Fields("tipo") = ""
   mytablex.Fields("local") = extra_loquesea(Local1) '
   mytablex.Fields("serie") = ""
   mytablex.Fields("numero") = ""
   mytablex.Fields("acu") = "W"
   mytablex.Fields("usuario") = extra_loquesea(cajero)
   mytablex.Fields("paridad") = 0
   mytablex.Fields("local1") = "" & rxconsulta.Fields("local")
   mytablex.Fields("tipo1") = "" & rxconsulta.Fields("tipo")
   mytablex.Fields("serie1") = "" & rxconsulta.Fields("serie")
   mytablex.Fields("numero1") = "" & rxconsulta.Fields("numero")
   mytablex.Fields("cuota") = "" & rxconsulta.Fields("cuota")
   mytablex.Fields("moneda") = "" & rxconsulta.Fields("Moneda")
   mytablex.Fields("total") = Val("" & rxconsulta.Fields("saldo"))
   mytablex.Fields("paga") = Val("" & rxconsulta.Fields("saldo"))
   mytablex.Fields("estado") = "2"
   mytablex.Update
   End If
   End If
   rxconsulta.MoveNext
   Loop
mytablex.Close
If sw = 0 Then
   MsgBox "No ha Seleccionado Ningun Dato ", 48, "Aviso"
   Exit Sub
End If
'-----------------------------
trecaja.Caption = "INGRESO DINERO"
trecaja.afecta = "C"
trecaja.acu = "W"
trecaja.cajero = gusuario
trecaja.caja = "00"
trecaja.turno = "1"
trecaja.fecha = Format(Now, "dd/mm/yyyy")
trecaja.dia = Format(Now, "dd/mm/yyyy")
trecaja.tipoclie.Enabled = False
trecaja.codigo.Enabled = False
trecaja.moneda.Enabled = False
trecaja.tipoclie.Enabled = False
trecaja.ch89343.Visible = True
trecaja.d7823.Visible = True
trecaja.vienede = "CXC"
trecaja.Show 1
End If

If acu = "C" Then
trecaja.Caption = "EGRESO DINERO"
trecaja.afecta = "P"
trecaja.Local1 = Trim("" & rxconsulta.Fields("local"))
trecaja.acu = "V"
trecaja.cajero = gusuario
trecaja.caja = "00"
trecaja.turno = "1"
trecaja.fecha = Format(Now, "dd/mm/yyyy")
trecaja.dia = Format(Now, "dd/mm/yyyy")
'trecaja.fecha.Enabled = False
trecaja.tipoclie.Enabled = False
trecaja.codigo.Enabled = False
trecaja.moneda.Enabled = False
trecaja.ch89343.Visible = True
trecaja.d7823.Visible = True
trecaja.vienede = "CXC"
trecaja.tipoclie = Trim("" & rxconsulta.Fields("tipoclie"))
trecaja.codigo = Trim("" & rxconsulta.Fields("codigo"))
trecaja.nombre = Trim("" & rxconsulta.Fields("nombre"))
trecaja.moneda = Trim("" & rxconsulta.Fields("moneda"))
trecaja.Show 1
End If
Exit Sub
cmd234_err:
MsgBox "Seleccione un registro ", 48, "Aviso"
Exit Sub


End Sub


Sub exporta_excel()
Dim v As Long
Dim h As Long
Dim i As Long
Dim sdx As Double
Dim sdx1 As Double
Dim sdx2 As Double

Dim Heading(16) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion
Dim buf As String
    Heading(1) = "Lo"
    Heading(2) = "Tp"
    Heading(3) = "Serie"
    Heading(4) = "Numero"
    Heading(5) = "Fecha"
    Heading(6) = "Fechav"
    Heading(7) = "Nombre"
    Heading(8) = "Tipo"
    Heading(9) = "Serie"
    Heading(10) = "Numero"
    Heading(11) = "Fecha"
    Heading(12) = "Total"
    Heading(13) = "Abono"
    Heading(14) = "Saldo"
    
    Heading(15) = "Observa"
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    
With objExcel.ActiveSheet
        
    For i = 1 To 15 Step 1
        .Cells(1, i) = Heading(i)
    Next i
       
        .columns("A").ColumnWidth = 5
        .columns("B").ColumnWidth = 5
        .columns("C").ColumnWidth = 5
        .columns("D").ColumnWidth = 10
        .columns("E").ColumnWidth = 10
        .columns("F").ColumnWidth = 10
        .columns("G").ColumnWidth = 30
        .columns("H").ColumnWidth = 5
        .columns("I").ColumnWidth = 5
        .columns("J").ColumnWidth = 10
        .columns("K").ColumnWidth = 10
        .columns("L").ColumnWidth = 10
        .columns("M").ColumnWidth = 10
        
End With

sdx = 0
sdx1 = 0
sdx2 = 0
    
v = 2
h = 1
     Do
     If rxconsulta.EOF Then Exit Do
            objExcel.ActiveSheet.Cells(v, h) = "'" & rxconsulta.Fields("local")
            objExcel.ActiveSheet.Cells(v, h + 1) = "'" & rxconsulta.Fields("tipo")
            objExcel.ActiveSheet.Cells(v, h + 2) = "'" & rxconsulta.Fields("serie")
            objExcel.ActiveSheet.Cells(v, h + 3) = "'" & rxconsulta.Fields("numero")
            objExcel.ActiveSheet.Cells(v, h + 4) = "'" & rxconsulta.Fields("fecha")
             objExcel.ActiveSheet.Cells(v, h + 5) = "'" & rxconsulta.Fields("fechav")
            objExcel.ActiveSheet.Cells(v, h + 6) = "'" & rxconsulta.Fields("nombre")
            
            objExcel.ActiveSheet.Cells(v, h + 11) = "" & rxconsulta.Fields("total")
            sdx = sdx + Val("" & rxconsulta.Fields("total"))
            objExcel.ActiveSheet.Cells(v, h + 12) = "" & rxconsulta.Fields("abono")
            sdx1 = sdx1 + Val("" & rxconsulta.Fields("abono"))
            objExcel.ActiveSheet.Cells(v, h + 13) = "" & rxconsulta.Fields("saldo")
            sdx2 = sdx2 + Val("" & rxconsulta.Fields("saldo"))
            
            buf = ""
      If "" & rxconsulta.Fields("grupo") = "C" Then
         buf = "CREDITO"
      End If
      If "" & rxconsulta.Fields("grupo") = "A" Then
         buf = "ADELANTO"
      End If
      If "" & rxconsulta.Fields("grupo") = "O" Then
         buf = "ORDENTRABAJO"
      End If
      If "" & rxconsulta.Fields("grupo") = "D" Then
      buf = "ADEL.BANCO"
      End If
      objExcel.ActiveSheet.Cells(v, h + 14) = "" & buf
      objExcel.ActiveSheet.Cells(v, h + 15) = "" & rxconsulta.Fields("observa")
      
            v = v + 1
            imprime_ecuenta v
     rxconsulta.MoveNext
     Loop
     
     objExcel.ActiveSheet.Cells(v, 12) = "" & sdx
     objExcel.ActiveSheet.Cells(v, 13) = "" & sdx1
     objExcel.ActiveSheet.Cells(v, 14) = "" & sdx2
 
 Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
 
 
End Sub
Sub imprime_ecuenta(v As Long)
Dim mytablex As New adodb.Recordset
Dim buf As String
Dim found As Integer
Dim sdx As Double
Dim sw As Integer
sdx = 0
sw = 0
If acu = "V" Then
mytablex.Open "select * from cuentacd where local='" & Trim(rxconsulta.Fields("local")) & "' and tipo1='" & Trim(rxconsulta.Fields("tipo")) & "' and serie1='" & Trim(rxconsulta.Fields("serie")) & "' and numero1='" & Trim(rxconsulta.Fields("numero")) & "'", cn, adOpenStatic, adLockOptimistic
End If
If acu = "C" Then
mytablex.Open "select * from cuentapd where local='" & Trim(rxconsulta.Fields("local")) & "' and tipo1='" & Trim(rxconsulta.Fields("tipo")) & "' and serie1='" & Trim(rxconsulta.Fields("serie")) & "' and numero1='" & Trim(rxconsulta.Fields("numero")) & "'", cn, adOpenStatic, adLockOptimistic
End If

If mytablex.RecordCount = 0 Then
   mytablex.Close
   Exit Sub
End If

Do
If mytablex.EOF Then Exit Do
      objExcel.ActiveSheet.Cells(v, 7) = "'Recibo"
      objExcel.ActiveSheet.Cells(v, 8) = "'" & mytablex.Fields("tipo")
      objExcel.ActiveSheet.Cells(v, 9) = "'" & mytablex.Fields("serie")
      objExcel.ActiveSheet.Cells(v, 10) = "'" & mytablex.Fields("numero")
      objExcel.ActiveSheet.Cells(v, 11) = "'" & mytablex.Fields("fecha")
      objExcel.ActiveSheet.Cells(v, 13) = "" & mytablex.Fields("paga")
      v = v + 1
   mytablex.MoveNext
Loop
mytablex.Close

End Sub



Sub cabecera_documento()
Dim buf As String
Dim i As Integer
Dim found As Integer
    If contlin > 0 Then
       buf = Chr$(12)
       found = formateaa(buf, Len(buf), 0, 0)
    End If
    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = "Reporte de Cuentas Corrientes x Cobrar  "
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("", 65, 0, 0)
    found = formateaa("-------RECIBO DE PAGO----------", 40, 2, 0)
    
    found = formateaa("Lo", 3, 0, 0)
    found = formateaa("Tp", 3, 0, 0)
    found = formateaa("Srie", 5, 0, 0)
    found = formateaa("Numero", 12, 0, 0)
    found = formateaa("Fecha", 11, 0, 0)
    found = formateaa("FechaV", 11, 0, 0)
    found = formateaa("Nombre", 20, 0, 0)
    found = formateaa("Tip Srie Numero ", 20, 0, 0)
    found = formateaa("Fecha", 11, 0, 0)
    found = formateaa("Total ", 11, 0, 1)
    found = formateaa("Abono ", 11, 0, 1)
    found = formateaa("Saldo", 11, 2, 1)
        
    buf = String(150, "-")
    found = formateaa(buf, 150, 2, 0)
    

End Sub
Sub cuerpo_programa_documento()
Dim buf As String
Dim found As Integer
Dim r As Long
On Error GoTo cmd788_err
suma5 = 0
suma6 = 0
suma7 = 0

Do
If rxconsulta.EOF Then Exit Do
      buf = "" & rxconsulta.Fields("LOCAL")
      found = formateaa(buf, 2, 0, 0)
      found = formateaa("", 1, 0, 0)
      buf = "" & rxconsulta.Fields("tipo")
      found = formateaa(buf, 2, 0, 0)
      found = formateaa("", 1, 0, 0)
      buf = "" & rxconsulta.Fields("serie")
      found = formateaa(buf, 4, 0, 0)
      found = formateaa("", 1, 0, 0)
      buf = "" & rxconsulta.Fields("numero")
      found = formateaa(buf, 11, 0, 0)
      found = formateaa("", 1, 0, 0)
      buf = "" & rxconsulta.Fields("fecha")
      found = formateaa(buf, 10, 0, 0)
      found = formateaa("", 1, 0, 0)
      buf = "" & rxconsulta.Fields("fechaV")
      found = formateaa(buf, 10, 0, 0)
      found = formateaa("", 1, 0, 0)
      buf = "" & rxconsulta.Fields("nombre")
      found = formateaa(buf, 19, 0, 0)
      found = formateaa("", 1, 0, 0)
      found = formateaa("", 23, 0, 0)
      found = formateaa("", 8, 0, 0)
      buf = Format(Val("" & rxconsulta.Fields("total")), "0.00")
      found = formateaa(buf, 10, 0, 1)
      found = formateaa("", 1, 0, 0)
      buf = Format(Val("" & rxconsulta.Fields("abono")), "0.00")
      found = formateaa(buf, 10, 0, 1)
      found = formateaa("", 1, 0, 0)
      buf = Format(Val("" & rxconsulta.Fields("saldo")), "0.00")
      found = formateaa(buf, 10, 0, 1)
      found = formateaa("", 1, 0, 0)
      buf = ""
      If "" & rxconsulta.Fields("grupo") = "C" Then
         buf = "CREDITO"
      End If
      If "" & rxconsulta.Fields("grupo") = "A" Then
         buf = "ADELANTO"
      End If
      If "" & rxconsulta.Fields("grupo") = "O" Then
         buf = "ORDENTRABAJO"
      End If
      If "" & rxconsulta.Fields("grupo") = "D" Then
      buf = "ADEL.BANCO"
      End If
      found = formateaa(buf, 10, 0, 0)
      found = formateaa("", 1, 2, 0)
      
      nlineas
      suma5 = suma5 + Val("" & rxconsulta.Fields("total"))
      suma6 = suma6 + Val("" & rxconsulta.Fields("abono"))
      suma7 = suma7 + Val("" & rxconsulta.Fields("saldo"))
      
      'If "" & rxconsulta.fields("grupo") = "O" Then
      '   found = imprime_orden_cruce()
      'End If
      'MsgBox ""
      found = imprime_cuentacd()
      If producto = "S" Then
          imprime_productos
      End If
      rxconsulta.MoveNext
      
Loop
      found = formateaa("", 96, 0, 0)
      buf = Format(suma5, "0.00")
      found = formateaa(buf, 10, 0, 1)
      found = formateaa("", 1, 0, 0)
      buf = Format(suma6, "0.00")
      found = formateaa(buf, 10, 0, 1)
      found = formateaa("", 1, 0, 0)
      buf = Format(suma7, "0.00")
      found = formateaa(buf, 10, 0, 1)
      found = formateaa("", 1, 2, 0)

Exit Sub
cmd788_err:
MsgBox "Aviso en cuerpo programa " + error$, 48, "Aviso"
Exit Sub
End Sub
Function imprime_cuentacd()
Dim mytablex As New adodb.Recordset
Dim buf As String
Dim found As Integer
Dim sdx As Double
Dim sw As Integer
sdx = 0
sw = 0
If acu = "V" Then
mytablex.Open "select * from cuentacd where local='" & Trim(rxconsulta.Fields("local")) & "' and tipo1='" & Trim(rxconsulta.Fields("tipo")) & "' and serie1='" & Trim(rxconsulta.Fields("serie")) & "' and numero1='" & Trim(rxconsulta.Fields("numero")) & "'", cn, adOpenStatic, adLockOptimistic
End If
If acu = "C" Then
mytablex.Open "select * from cuentapd where local='" & Trim(rxconsulta.Fields("local")) & "' and tipo1='" & Trim(rxconsulta.Fields("tipo")) & "' and serie1='" & Trim(rxconsulta.Fields("serie")) & "' and numero1='" & Trim(rxconsulta.Fields("numero")) & "'", cn, adOpenStatic, adLockOptimistic
End If

If mytablex.RecordCount = 0 Then
   mytablex.Close
   Exit Function
End If

Do
If mytablex.EOF Then Exit Do
      sw = 1
      found = formateaa("", 65, 0, 0)
      buf = "" & mytablex.Fields("tipo")
      found = formateaa(buf, 3, 0, 0)
      found = formateaa("", 1, 0, 0)
      buf = "" & mytablex.Fields("serie")
      found = formateaa(buf, 4, 0, 0)
      found = formateaa("", 1, 0, 0)
      buf = "" & mytablex.Fields("numero")
      found = formateaa(buf, 11, 0, 0)
      found = formateaa("", 1, 0, 0)
      buf = "" & mytablex.Fields("fecha")
      found = formateaa(buf, 10, 0, 0)
      found = formateaa("", 1, 0, 0)
      found = formateaa("", 10, 0, 0)
      buf = Format(Val("" & mytablex.Fields("paga")), "0.00")
      found = formateaa(buf, 10, 0, 1)
      found = formateaa("", 1, 2, 0)
      nlineas
   mytablex.MoveNext
Loop
mytablex.Close

End Function
Function imprime_orden_cruce()
Dim mytablex As New adodb.Recordset
Dim buf As String
Dim found As Integer
Dim sdx As Double
Dim sw As Integer
sdx = 0
sw = 0

mytablex.Open "select * from factura where local='" & rxconsulta.Fields("local") & "' and tipo='" & rxconsulta.Fields("tipo") & "' and serie='" & rxconsulta.Fields("serie") & "' and numero='" & rxconsulta.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount = 0 Then
   mytablex.Close
   Exit Function
End If

Do
If mytablex.EOF Then Exit Do
      sw = 1
      found = formateaa("", 65, 0, 0)
      buf = "" & mytablex.Fields("tipo")
      found = formateaa(buf, 3, 0, 0)
      found = formateaa("", 1, 0, 0)
      buf = "" & mytablex.Fields("serie")
      found = formateaa(buf, 4, 0, 0)
      found = formateaa("", 1, 0, 0)
      buf = "" & mytablex.Fields("numero")
      found = formateaa(buf, 11, 0, 0)
      found = formateaa("", 1, 0, 0)
      buf = "" & mytablex.Fields("fecha")
      found = formateaa(buf, 10, 0, 0)
      found = formateaa("", 1, 0, 0)
      found = formateaa("", 10, 0, 0)
      buf = Format(Val("" & mytablex.Fields("Total")), "0.00")
      found = formateaa(buf, 10, 0, 1)
      found = formateaa("", 1, 2, 0)
      nlineas
   mytablex.MoveNext
Loop
mytablex.Close

End Function



Sub nlineas()
    contlin = contlin + 1
    If contlin > 45 Then
       cabecera_documento
    End If
End Sub



Private Sub dkj8923_Click()
End Sub

Private Sub dl993_Click()
exporta_excel
End Sub

Private Sub dlo23211_Click()
If Frame1.Visible = True Then
   Frame1.Visible = False
   Exit Sub
End If
If Frame2.Visible = True Then
   Frame2.Visible = False
   dbgrid2.SetFocus
   Exit Sub
End If


texplcxc.Hide
Unload texplcxc
End Sub

Private Sub Form_Activate()
If acu = "V" Then
   xnameclie = "clientes"
   xcuentaco = "cuentac"
End If
If acu = "C" Then
   xnameclie = "proveedo"
   xcuentaco = "cuentap"
End If
'fechai = "01/01/" & Format(Year(Now), "0000")
fechai = Format(Now, "dd/mm/yyyy")
fechaf = Format(Now, "dd/mm/yyyy")
carga_inicial
sql_letras
End Sub
Sub carga_inicial()
Dim mytablex As New adodb.Recordset
Local1.Clear
Local1.AddItem "%"

vendedor.Clear
vendedor.AddItem "%"
cajero.Clear
cajero.AddItem "%"

mytablex.Open "select * from vendedor ", cn, adOpenStatic, adLockOptimistic


Do
If mytablex.EOF Then Exit Do
vendedor.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
cajero.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
mytablex.MoveNext
Loop
mytablex.Close
cajero.ListIndex = 0
vendedor.ListIndex = 0



mytablex.Open "select * from tlocal ", cn, adOpenStatic, adLockOptimistic
Do
If mytablex.EOF Then Exit Do
Local1.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
mytablex.MoveNext
Loop
mytablex.Close
Local1.ListIndex = 0
If Local1.ListCount = 2 Then
Local1.ListIndex = 1
End If



caja.Clear
caja.AddItem "%"
mytablex.Open "select * from parameca", cn, adOpenStatic, adLockOptimistic
Do
If mytablex.EOF Then Exit Do
caja.AddItem "" & mytablex.Fields("caja") & "|" & mytablex.Fields("DESCRIPCIO")

mytablex.MoveNext
Loop
mytablex.Close
caja.ListIndex = 0

turno.Clear
turno.AddItem "%"
mytablex.Open "select * from turno", cn, adOpenStatic, adLockOptimistic
Do
If mytablex.EOF Then Exit Do
turno.AddItem "" & mytablex.Fields("turno") & "|" & mytablex.Fields("DESCRIPCIO")
mytablex.MoveNext
Loop
mytablex.Close
turno.ListIndex = 0


tipo.Clear
tipo.AddItem "%"
mytablex.Open "select * from tipo", cn, adOpenStatic, adLockOptimistic
Do
If mytablex.EOF Then Exit Do
tipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("DESCRIPCIO")
mytablex.MoveNext
Loop
mytablex.Close
tipo.ListIndex = 0
End Sub


Sub sql_letras()
On Error GoTo cmd37_err
Dim buf As String

If Not IsDate(fechai) Then Exit Sub
If Not IsDate(fechaf) Then Exit Sub
buf = "select * from " & xcuentaco & " where "
buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
If Local1 <> "%" Then
buf = buf & " and local like '" & extra_loquesea(Local1) & "'"
End If
If cajero <> "%" Then
buf = buf & " and usuario like '" & extra_loquesea(cajero) & "'"
End If
If vendedor <> "%" Then
   buf = buf & " and vendedor like '" & extra_loquesea(vendedor) & "'"
End If
If caja <> "%" Then
buf = buf & " and caja like '" & extra_loquesea(caja) & "'"
End If
If turno <> "%" Then
buf = buf & " and turno like '" & extra_loquesea(turno) & "'"
End If
If tipo <> "%" Then
buf = buf & " and tipo like '" & extra_loquesea(tipo) & "'"
End If
If codigo <> "%" Then
buf = buf & " and codigo like '" & codigo & "'"
End If
If nombre <> "%" Then
buf = buf & " and nombre like '" & nombre & "'"
End If
If condicion <> "%" Then
   buf = buf & " AND " & condicion
End If

If servicios = "CREDITO" Then
buf = buf & " and grupo='C'"
End If
If servicios = "ANTICIPO DINERO" Then
buf = buf & " and grupo='A'"
End If
If servicios = "DEPOSITO BANCO" Then
buf = buf & " and grupo='D'"
End If
If servicios = "ORDEN TRABAJO" Then
buf = buf & " and grupo='O'"
End If

buf = buf & " order by grupo," & ordenado & " ,numero"
'MsgBox buf
   If rxconsulta.State = 1 Then rxconsulta.Close
   rxconsulta.Open buf, cn, adOpenStatic, adLockOptimistic
   If rxconsulta.EOF = True And rxconsulta.BOF = True Then
   End If
   Set dbgrid2.DataSource = rxconsulta
               suma_sql rxconsulta
               'dbgrid2.SetFocus
Exit Sub
cmd37_err:
MsgBox "Error en select " & error$, 48, "Aviso"
Exit Sub
End Sub
Sub suma_sql(mytablex As adodb.Recordset)
Dim xx As String
Dim xtotalh As Double
Dim xabonoh As Double
Dim xsaldoh As Double


Dim xtotalc As Double
Dim xabonoc As Double
Dim xsaldoc As Double

Dim xtotalo As Double
Dim xabonoo As Double
Dim xsaldoo As Double

Dim xtotala As Double
Dim xabonoa As Double
Dim xsaldoa As Double

xtotalc = 0
xabonoc = 0
xsaldoc = 0

xtotalo = 0
xabonoo = 0
xsaldoo = 0

xtotala = 0
xabonoa = 0
xsaldoa = 0

xtotalh = 0
xabonoh = 0
xsaldoh = 0




Do
If mytablex.EOF Then Exit Do

If "" & mytablex.Fields("grupo") = "O" Then
     xtotalo = xtotalo + Val("" & mytablex.Fields("total"))
     xabonoo = xabonoo + Val("" & mytablex.Fields("abono"))
     xsaldoo = xsaldoo + Val("" & mytablex.Fields("saldo"))
     GoTo amix
End If
If "" & mytablex.Fields("grupo") = "C" Then
     xtotalc = xtotalc + Val("" & mytablex.Fields("total"))
     xabonoc = xabonoc + Val("" & mytablex.Fields("abono"))
     xsaldoc = xsaldoc + Val("" & mytablex.Fields("saldo"))
     GoTo amix
End If
If "" & mytablex.Fields("grupo") = "A" Then  'adelantos
     xtotala = xtotala + Val("" & mytablex.Fields("total"))
     xabonoa = xabonoa + Val("" & mytablex.Fields("abono"))
     xsaldoa = xsaldoa + Val("" & mytablex.Fields("saldo"))
     GoTo amix
End If
If "" & mytablex.Fields("grupo") = "D" Then  'depositos bancos
     xtotalh = xtotalh + Val("" & mytablex.Fields("total"))
     xabonoh = xabonoh + Val("" & mytablex.Fields("abono"))
     xsaldoh = xsaldoh + Val("" & mytablex.Fields("saldo"))
     GoTo amix
End If

amix:
mytablex.MoveNext
Loop

totalc = Format(xtotalc, "0.00")
abonoc = Format(xabonoc, "0.00")
saldoc = Format(xsaldoc, "0.00")

totala = Format(xtotala, "0.00")
abonoa = Format(xabonoa, "0.00")
saldoa = Format(xsaldoa, "0.00")

totalo = Format(xtotalo, "0.00")
abonoo = Format(xabonoo, "0.00")
saldoo = Format(xsaldoo, "0.00")

totalh = Format(xtotalh, "0.00")
abonoh = Format(xabonoh, "0.00")
saldoh = Format(xsaldoh, "0.00")


End Sub

Private Sub Form_Load()


servicios.Clear
servicios.AddItem "CREDITO"
servicios.AddItem "ANTICIPO DINERO"
servicios.AddItem "DEPOSITO BANCO"
servicios.AddItem "ORDEN TRABAJO"
servicios.AddItem "%"

servicios.ListIndex = 0

producto.Clear
producto.AddItem "N"
producto.AddItem "S"
producto.ListIndex = 0
ordenado.Clear
ordenado.AddItem "fecha"
ordenado.AddItem "Codigo"
ordenado.AddItem "fechaV"
ordenado.AddItem "vendedor"
ordenado.AddItem "tipo"
'ordenado.AddItem "STR(numero)"
ordenado.AddItem "Usuario"
ordenado.AddItem "caja"
ordenado.AddItem "turno"
ordenado.AddItem "nombre"
ordenado.ListIndex = 0
condicion.Clear
condicion.AddItem "%"
condicion.AddItem "Saldo>0"
condicion.AddItem "Saldo>=0"
condicion.AddItem "Saldo<0"
condicion.AddItem "Saldo<=0"
condicion.AddItem "Saldo=0"
condicion.ListIndex = 0
End Sub

Private Sub KU7zom8_Click()
Dim mytablex As New adodb.Recordset
On Error GoTo cmd87999_err
Dim sdx As Double
Dim sdx1 As Double
If Frame1.Visible = True Then Exit Sub
sdx = 0
sdx1 = 0
tsdx = ""
tsdx1 = ""
hlocal = Trim(rxconsulta.Fields("local"))
htipo = Trim(rxconsulta.Fields("tipo"))
hserie = Trim(rxconsulta.Fields("serie"))
hnumero = Trim(rxconsulta.Fields("numero"))
hmoneda = Trim(rxconsulta.Fields("moneda"))
htotal = Trim(rxconsulta.Fields("Total"))
hsaldo = Trim(rxconsulta.Fields("saldo"))
If acu = "V" Then
mytablex.Open "select Local,Tipo,Serie,Numero,Moneda as M,Paga,Fecha,Usuario,Caja,Turno,Hora,Estado as E from cuentacd where local='" & Trim(rxconsulta.Fields("local")) & "' and tipo1='" & Trim(rxconsulta.Fields("tipo")) & "' and serie1='" & Trim(rxconsulta.Fields("serie")) & "' and numero1='" & Trim(rxconsulta.Fields("numero")) & "'", cn, adOpenStatic, adLockOptimistic
End If
If acu = "C" Then
mytablex.Open "select Local,Tipo,Serie,Numero,Moneda as M,Paga,Fecha,Usuario,Caja,Turno,Hora,Estado as E from cuentapd where local='" & Trim(rxconsulta.Fields("local")) & "' and tipo1='" & Trim(rxconsulta.Fields("tipo")) & "' and serie1='" & Trim(rxconsulta.Fields("serie")) & "' and numero1='" & Trim(rxconsulta.Fields("numero")) & "'", cn, adOpenStatic, adLockOptimistic
End If
   If mytablex.RecordCount = 0 Then
      mytablex.Close
      Exit Sub
   End If
   Frame1.Visible = True
   Set dbgrid3.DataSource = mytablex
   
   Do
   If mytablex.EOF Then Exit Do
   If "" & mytablex.Fields("m") = "S" Then
      sdx = sdx + Val("" & mytablex.Fields("paga"))
   End If
   If "" & mytablex.Fields("m") = "D" Then
      sdx1 = sdx1 + Val("" & mytablex.Fields("paga"))
   End If
   mytablex.MoveNext
   Loop
   tsdx = "" & sdx
   tsdx1 = "" & sdx1
   Exit Sub
cmd87999_err:
   MsgBox "Seleccione un dato ", 48, "Aviso"
   Exit Sub
   
   
End Sub

Private Sub Label25_Click()
Dim mytablex As New adodb.Recordset
Dim found As Integer
Dim sdx As Double
On Error GoTo cmd12566_err
found = valida_factvta()
If found = 1 Then
   MsgBox "No se puede borrar,Viene de facturacion..", 48, "Aviso"
   Exit Sub
End If
If MsgBox("Desea borra el Recibo " + dbgrid3.columns("numero"), 1, "Aviso") <> 1 Then Exit Sub
If acu = "V" Then
mytablex.Open "select * from cuentacd where local='" & dbgrid3.columns("local") & "' and tipo='" & dbgrid3.columns("tipo") & "' and serie='" & dbgrid3.columns("serie") & "' and numero='" & dbgrid3.columns("numero") & "'", cn, adOpenStatic, adLockOptimistic
End If
If acu = "C" Then
mytablex.Open "select * from cuentapd where local='" & dbgrid3.columns("local") & "' and tipo='" & dbgrid3.columns("tipo") & "' and serie='" & dbgrid3.columns("serie") & "' and numero='" & dbgrid3.columns("numero") & "'", cn, adOpenStatic, adLockOptimistic
End If

If mytablex.RecordCount = 0 Then
   mytablex.Close
   Exit Sub
End If
sdx = 0
Do
If mytablex.EOF Then Exit Do
sdx = sdx + Val("" & mytablex.Fields("paga"))
mytablex.MoveNext
Loop
mytablex.Close
'MsgBox ""
'rxconsulta.Edit
rxconsulta.Fields("abono") = Val("" & rxconsulta.Fields("abono")) - sdx
rxconsulta.Fields("saldo") = Val("" & rxconsulta.Fields("total")) + Val("" & rxconsulta.Fields("interes")) - Val("" & rxconsulta.Fields("abono"))
rxconsulta.Update

cn.Execute ("delete from recibo where local='" & dbgrid3.columns("local") & "' and tipo='" & dbgrid3.columns("tipo") & "' and serie='" & dbgrid3.columns("serie") & "' and numero='" & dbgrid3.columns("numero") & "'")
cn.Execute ("delete from fpagov where local='" & dbgrid3.columns("local") & "' and tipo='" & dbgrid3.columns("tipo") & "' and serie='" & dbgrid3.columns("serie") & "' and numero='" & dbgrid3.columns("numero") & "'")
cn.Execute ("delete from cuentapd where local='" & dbgrid3.columns("local") & "' and tipo='" & dbgrid3.columns("tipo") & "' and serie='" & dbgrid3.columns("serie") & "' and numero='" & dbgrid3.columns("numero") & "'")
Label29_Click
sql_letras
'sql_letras
Exit Sub
cmd12566_err:
MsgBox "Seleccione un documento ", 48, "Aviso"
Exit Sub

End Sub
Function valida_factvta()
Dim mytablex As New adodb.Recordset
mytablex.Open "select * from factura where local='" & Trim(dbgrid3.columns("local")) & "' and tipo='" & Trim(dbgrid3.columns("tipo")) & "' and serie='" & Trim(dbgrid3.columns("serie")) & "' and numero='" & Trim(dbgrid3.columns("numero")) & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   valida_factvta = 1
End If
mytablex.Close

End Function

Private Sub Label29_Click()
Frame1.Visible = False
End Sub

Private Sub ldo232_Click()
If Frame1.Visible = True Then Exit Sub
If Local1 = "%" Then
   MsgBox "Seleccione local ", 48, "Aviso"
   Exit Sub
End If
If Frame2.Visible = True Then Exit Sub
tincxc.Local1 = extra_loquesea(Local1)
tincxc.usuario = gusuario
tincxc.caja = "00"
tincxc.turno = "1"
tincxc.bandera = "NUEVO"
tincxc.acu = acu
tincxc.Show 1
End Sub

Private Sub lk993_Click()
If Frame1.Visible = True Then Exit Sub
Frame2.Visible = True


End Sub

Private Sub mofdi782_Click()
Dim found As Integer
On Error GoTo cmd345_err
If Frame1.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub

'If Frame1.Visible = True Then Exit Sub
If Val("" & rxconsulta.Fields("saldo")) <= 0 Then
   MsgBox "Documento ya cancelado ", 48, "Aviso"
   Exit Sub
End If

found = valida_factura()
If found = 1 Then
   MsgBox "No se puede Modificar viene de documento ", 48, "Aviso"
   Exit Sub
End If

tincxc.bandera = "MODIFICA"
tincxc.acu = acu
tincxc.Local1.Enabled = False
tincxc.tipo.Enabled = False
tincxc.serie.Enabled = False
tincxc.numero.Enabled = False
tincxc.cuota.Enabled = False
tincxc.tipoclie.Enabled = False
tincxc.codigo.Enabled = False



tincxc.anticipo = "" & rxconsulta.Fields("anticipo")
tincxc.Local1 = "" & rxconsulta.Fields("local")

tincxc.tipo = "" & rxconsulta.Fields("tipo")
tincxc.usuario = "" & rxconsulta.Fields("usuario")

tincxc.caja = "" & rxconsulta.Fields("caja")
tincxc.turno = "" & rxconsulta.Fields("turno")
tincxc.grupo = "" & rxconsulta.Fields("grupo")



tincxc.serie = "" & rxconsulta.Fields("serie")
tincxc.numero = "" & rxconsulta.Fields("numero")
tincxc.cuota = "" & rxconsulta.Fields("cuota")
tincxc.tipoclie = "" & rxconsulta.Fields("tipoclie")
tincxc.codigo = "" & rxconsulta.Fields("codigo")
tincxc.nombre = "" & rxconsulta.Fields("nombre")
tincxc.zona = "" & rxconsulta.Fields("zona")

tincxc.vendedor = "" & rxconsulta.Fields("vendedor")
tincxc.moneda = "" & rxconsulta.Fields("moneda")
tincxc.total = "" & rxconsulta.Fields("total")
tincxc.interes = "" & rxconsulta.Fields("interes")
tincxc.abono = "" & rxconsulta.Fields("abono")
tincxc.saldo = "" & rxconsulta.Fields("saldo")
tincxc.fecha = "" & rxconsulta.Fields("fecha")
tincxc.fechav = "" & rxconsulta.Fields("fechav")
tincxc.Show 1
Exit Sub
cmd345_err:
MsgBox "Seleccione un registro ", 48, "Aviso"
Exit Sub

End Sub
Function busca_fpago(buf As String) As String
Dim mytablex As New adodb.Recordset
mytablex.Open "select * from fpago where fpago='" & buf & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   busca_fpago = "" & mytablex.Fields("tipo")
End If
'------------------------------------- ------------
mytablex.Close
End Function

Function busca_factura() As Double
Dim mytablex As New adodb.Recordset
mytablex.Open "select * from factura where local='" & rxconsulta.Fields("local") & "' and tipo1='" & rxconsulta.Fields("tipo") & "' and serie1='" & rxconsulta.Fields("serie") & "' and numero1='" & rxconsulta.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   busca_factura = Val("" & mytablex.Fields("total"))
End If
mytablex.Close
End Function
Sub imprime_productos()
Dim mytablex As New adodb.Recordset
Dim sw As Integer
Dim buf As String
sw = 0
mytablex.Open "select * from detalle where local='" & rxconsulta.Fields("local") & "' and tipo='" & rxconsulta.Fields("tipo") & "' and serie='" & rxconsulta.Fields("serie") & "' and numero='" & rxconsulta.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   Do
   If mytablex.EOF Then Exit Do
   
      '-----------------------------------------------
      found = formateaa(">>>>>>>", 7, 0, 0)
       buf = "" & mytablex.Fields("producto")
      found = formateaa(buf, 7, 0, 0)
      found = formateaa("", 1, 0, 0)
       buf = "" & mytablex.Fields("descripcio")
      found = formateaa(buf, 30, 0, 0)
      found = formateaa("", 1, 0, 0)
       buf = "" & mytablex.Fields("unidad")
      found = formateaa(buf, 3, 0, 0)
      found = formateaa("", 1, 0, 0)
       buf = "" & mytablex.Fields("factor")
      found = formateaa(buf, 4, 0, 0)
      found = formateaa("", 1, 0, 0)
       buf = "" & mytablex.Fields("cantidad")
      found = formateaa(buf, 7, 0, 1)
      found = formateaa("", 1, 0, 0)
       buf = "" & mytablex.Fields("precio")
      found = formateaa(buf, 7, 0, 1)
      found = formateaa("", 1, 0, 0)
       buf = "" & mytablex.Fields("total")
      found = formateaa(buf, 9, 0, 1)
      found = formateaa("", 1, 2, 0)
      nlineas
      
      '-----------------------------------------------
   
   mytablex.MoveNext
   Loop
End If
mytablex.Close
End Sub

Private Sub reo0922_Click()

End Sub
Function valida_factura()
Dim mytablex As New adodb.Recordset
mytablex.Open "select * from factura where local='" & Trim(rxconsulta.Fields("local")) & "' and tipo='" & Trim(rxconsulta.Fields("tipo")) & "' and serie='" & Trim(rxconsulta.Fields("serie")) & "' and numero='" & Trim(rxconsulta.Fields("numero")) & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   valida_factura = 1
End If
mytablex.Close

End Function
Function valida_recibo()
Dim mytablex As New adodb.Recordset
If acu = "V" Then
mytablex.Open "select local from cuentacd where local='" & Trim(rxconsulta.Fields("local")) & "' and tipo1='" & Trim(rxconsulta.Fields("tipo")) & "' and serie1='" & Trim(rxconsulta.Fields("serie")) & "' and numero1='" & Trim(rxconsulta.Fields("numero")) & "'", cn, adOpenStatic, adLockOptimistic
End If
If acu = "C" Then
mytablex.Open "select Local from cuentapd where local='" & Trim(rxconsulta.Fields("local")) & "' and tipo1='" & Trim(rxconsulta.Fields("tipo")) & "' and serie1='" & Trim(rxconsulta.Fields("serie")) & "' and numero1='" & Trim(rxconsulta.Fields("numero")) & "'", cn, adOpenStatic, adLockOptimistic
End If
If mytablex.RecordCount > 0 Then
   valida_recibo = 1
End If
mytablex.Close
End Function

