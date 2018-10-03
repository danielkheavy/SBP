VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ttipodoc 
   BackColor       =   &H00FFFF00&
   Caption         =   "Tabla de Tipos de Documentos"
   ClientHeight    =   8385
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   13440
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   0
      TabIndex        =   65
      Top             =   0
      Visible         =   0   'False
      Width           =   11895
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
         TabIndex        =   68
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
         TabIndex        =   67
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
         Left            =   8280
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   6735
         Left            =   240
         TabIndex        =   69
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
   Begin VB.TextBox cuenta7 
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
      Left            =   9480
      MaxLength       =   8
      TabIndex        =   63
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox cuenta6 
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
      Left            =   9480
      MaxLength       =   8
      TabIndex        =   61
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox cuenta5 
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
      Left            =   9480
      MaxLength       =   8
      TabIndex        =   59
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox cuenta4 
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
      Left            =   9480
      MaxLength       =   8
      TabIndex        =   57
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox archivoe 
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
      Left            =   4680
      MaxLength       =   30
      TabIndex        =   56
      Top             =   6360
      Width           =   2415
   End
   Begin VB.TextBox cajachica 
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
      Left            =   7080
      MaxLength       =   1
      TabIndex        =   54
      Top             =   7200
      Width           =   375
   End
   Begin VB.TextBox anticipo 
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
      Left            =   5040
      MaxLength       =   1
      TabIndex        =   52
      Top             =   7200
      Width           =   615
   End
   Begin VB.TextBox obliga 
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
      TabIndex        =   50
      Top             =   7200
      Width           =   615
   End
   Begin VB.TextBox crucedoc 
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
      TabIndex        =   48
      Top             =   6840
      Width           =   615
   End
   Begin VB.TextBox cuenta3 
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
      Left            =   9480
      MaxLength       =   8
      TabIndex        =   45
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox cuenta2 
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
      Left            =   9480
      MaxLength       =   8
      TabIndex        =   43
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox cuenta1 
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
      Left            =   9480
      MaxLength       =   8
      TabIndex        =   41
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox ts 
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
      Left            =   10200
      MaxLength       =   11
      TabIndex        =   40
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox te 
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
      Left            =   10200
      MaxLength       =   11
      TabIndex        =   38
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox archivo 
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
      TabIndex        =   34
      Top             =   6360
      Width           =   2415
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
      TabIndex        =   33
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox tipodoc 
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
      TabIndex        =   32
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox nrolineas 
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
      TabIndex        =   30
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox contable 
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
      MaxLength       =   20
      TabIndex        =   28
      Top             =   6000
      Width           =   2415
   End
   Begin VB.TextBox sunat 
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
      TabIndex        =   26
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox puerto 
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
      TabIndex        =   23
      Top             =   5280
      Width           =   4335
   End
   Begin VB.TextBox numero 
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
      Left            =   4680
      MaxLength       =   11
      TabIndex        =   21
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox serie 
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
      TabIndex        =   19
      Top             =   4560
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
      Top             =   1320
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
      MaxLength       =   3
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
      Picture         =   "TTIPODOC.frx":0000
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
      Picture         =   "TTIPODOC.frx":1212
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
      Picture         =   "TTIPODOC.frx":2424
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
      Picture         =   "TTIPODOC.frx":3636
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
      Picture         =   "TTIPODOC.frx":4848
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
      Height          =   645
      Left            =   720
      MaskColor       =   &H00E0E0E0&
      Picture         =   "TTIPODOC.frx":5A5A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Borrar registro"
      Top             =   0
      Width           =   720
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "TTIPODOC.frx":6C6C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Consulta"
      Top             =   0
      Width           =   735
   End
   Begin VB.Label local1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   6120
      TabIndex        =   70
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label31 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   7320
      TabIndex        =   64
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label30 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   7320
      TabIndex        =   62
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label29 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   7320
      TabIndex        =   60
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label28 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   7320
      TabIndex        =   58
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label27 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rubro "
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
      TabIndex        =   55
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Label Label26 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Anticipo/BcoDepo"
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
      Left            =   3000
      TabIndex        =   53
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Label Label25 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cruce Obligatorio"
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
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Label Label24 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hacer Cruce Doc."
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
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cuentas Contable"
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
      Left            =   7320
      TabIndex        =   47
      Top             =   2760
      Width           =   3855
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
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
      Left            =   7320
      TabIndex        =   46
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Impuesto"
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
      Left            =   7320
      TabIndex        =   44
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Subtotal"
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
      Left            =   7320
      TabIndex        =   42
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Traslado Tipodoc "
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
      Left            =   7920
      TabIndex        =   39
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label flage 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   6960
      TabIndex        =   37
      Top             =   840
      Width           =   495
   End
   Begin VB.Label grupo 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   7560
      TabIndex        =   36
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Archivo Formato"
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
      TabIndex        =   35
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NumeroLineas"
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
      TabIndex        =   31
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CuentaContable"
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
      TabIndex        =   29
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RelacionSunat"
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
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AlmacenBase"
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
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PuertoImpresion"
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
      TabIndex        =   24
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Numero"
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
      Left            =   3600
      TabIndex        =   22
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Serie"
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
      TabIndex        =   20
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bancos  :X.Cargos Y.Descargos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3960
      Width           =   6015
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Recibos  :V.ReciboEgreso W.ReciboIngreso"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3720
      Width           =   6015
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Produccion : U.OrdenProduccion  Z.Traslado"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   6015
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Salidas  :T.GuiaRemision"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Width           =   6015
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Entradas:S.GuiaRemision"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   6015
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"TTIPODOC.frx":7E7E
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   6015
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"TTIPODOC.frx":7F15
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   6015
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Grupo"
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
      TabIndex        =   11
      Top             =   1680
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
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TipoDocumento"
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
Attribute VB_Name = "ttipodoc"
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

Private Sub anticipo_Change()
'If tipodoc <> "V" And tipodoc <> "W" Then
'   anticipo = ""
'End If
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
contable.SetFocus

End Sub

Private Sub bodega_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   puerto.SetFocus
   Exit Sub
End If

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
   cad = "SELECT * FROM tipo "
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True Or rconsulta.BOF = True Then
      rconsulta.Close
      Exit Sub
   End If
Frame1.Visible = True
Frame1.Enabled = True
opcion1 = "1"
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
      cad = "SELECT Descripcio,Tipo from tipo "
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT Descripcio,Tipo from Tipo   where " & Combo1 & " like '" & buffer & "%'"
   End If
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      rconsulta.Close
      buffer.SetFocus
      Exit Sub
   End If
   Set dbGrid1.DataSource = rconsulta
   dbGrid1.Columns(0).Width = 4000
   dbGrid1.Columns(1).Width = 2000
   
   
   If sw = 1 Then
      dbGrid1.SetFocus
   End If
   Exit Sub
End If



End Sub



Private Sub contable_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   bodega.SetFocus
   Exit Sub
End If

End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
  
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
   codigo = dbGrid1.Columns(1)
   Frame1.Visible = False
   Frame1.Enabled = False
   codigo.SetFocus
   codigo_KeyPress 13
End If
End Sub


Private Sub descripcio_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
tipodoc.SetFocus

End Sub



Private Sub descripcio_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   codigo.SetFocus
   Exit Sub
End If

End Sub

Private Sub djuer1_Click()
If Frame1.Visible = True Then Exit Sub
reporgen.NAMETABLA = "tipo"
reporgen.Show 1

End Sub

Private Sub dlo132_Click()
If Frame1.Visible = True Then
   Frame1.Visible = False
   Exit Sub
End If
ttipodoc.Hide
Unload ttipodoc
End Sub



Private Sub Form_Load()
local1 = glocal
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "tipo"
Combo1.ListIndex = 0
End Sub
Sub inicializa()
archivoe = ""
cajachica = ""
anticipo = ""
obliga = ""
crucedoc = ""
cuenta1 = ""
cuenta2 = ""
cuenta3 = ""
cuenta4 = ""
cuenta5 = ""
cuenta6 = ""
cuenta7 = ""
te = ""
ts = ""
flage = ""
grupo = ""
descripcio = ""
tipodoc = ""
sunat = ""
serie = ""
numero = ""
nrolineas = ""
puerto = ""
bodega = ""
contable = ""
archivo = ""
End Sub
Function borra_registro()

On Error GoTo cmd56_err

cn.Execute ("DELETE   FROM tipo WHERE  tipo='" & Trim(codigo) & "'")
borra_registro = 1
Exit Function
cmd56_err:
MsgBox "Aviso en borra " + error$, 48, "Aviso"
Exit Function




End Function
Function busca_registro()

Dim rsexiste As New ADODB.Recordset
   rsexiste.Open "SELECT * FROM tipo where   tipo='" & Trim(codigo) & "'", cn, adOpenKeyset, adLockOptimistic
   If rsexiste.RecordCount > 0 Then  'si existe
      pone_registro rsexiste
      busca_registro = 1
   End If
 
End Function
Sub pone_registro(mytablex As ADODB.Recordset)
cajachica = "" & mytablex.Fields("cajachica")
anticipo = "" & mytablex.Fields("anticipo")
obliga = "" & mytablex.Fields("obliga")
crucedoc = "" & mytablex.Fields("crucedoc")
cuenta1 = "" & mytablex.Fields("cuenta1")
cuenta2 = "" & mytablex.Fields("cuenta2")
cuenta3 = "" & mytablex.Fields("cuenta3")
cuenta4 = "" & mytablex.Fields("cuenta4")
cuenta5 = "" & mytablex.Fields("cuenta5")
cuenta6 = "" & mytablex.Fields("cuenta6")
cuenta7 = "" & mytablex.Fields("cuenta7")


flage = "" & mytablex.Fields("flage")
grupo = "" & mytablex.Fields("grupo")
codigo = "" & mytablex.Fields("tipo")
descripcio = "" & mytablex.Fields("descripcio")
tipodoc = "" & mytablex.Fields("tipodoc")
sunat = "" & mytablex.Fields("sunat")
serie = "" & mytablex.Fields("serie")
numero = "" & mytablex.Fields("numero")
nrolineas = "" & mytablex.Fields("nrolineas")
puerto = "" & mytablex.Fields("puerto")
bodega = "" & mytablex.Fields("bodega")
contable = "" & mytablex.Fields("contable")
archivo = "" & mytablex.Fields("archivo")
archivoe = "" & mytablex.Fields("archivoe")
te = "" & mytablex.Fields("te")
ts = "" & mytablex.Fields("ts")

End Sub
Sub grabando(sw As Integer)


grupo = tipodoc
Select Case tipodoc
       Case "A", "B", "C", "D", "G", "F", "N", "T"
            flage = "S"
       Case "E", "J", "K", "L", "M", "P", "O", "S"
            flage = "E"
       
End Select
           
Select Case tipodoc
       Case "A", "B", "C", "D", "G"
            grupo = "V"
       Case "J", "K", "L", "M", "P"
            grupo = "C"
End Select

Dim cad As String


If sw = 0 Then
   cad = "INSERT INTO tipo VALUES('" & Trim(codigo) & "','"
   cad = cad & Trim(descripcio) & "','"
   cad = cad & Trim(tipodoc) & "','"
   cad = cad & Trim(sunat) & "','"
   cad = cad & Trim(serie) & "','"
   cad = cad & Trim(numero) & "','"
   cad = cad & Trim(nrolineas) & "','"
   cad = cad & Trim(puerto) & "','"
   cad = cad & Trim(bodega) & "','"
   cad = cad & Trim(contable) & "','"
   cad = cad & Trim(archivo) & "','"
   cad = cad & Trim(grupo) & "','"
   cad = cad & Trim(flage) & "','"
   cad = cad & Trim(te) & "','"
   cad = cad & Trim(ts) & "','"
   cad = cad & Trim(cuenta1) & "','"
   cad = cad & Trim(cuenta2) & "','"
   cad = cad & Trim(cuenta3) & "','"
   cad = cad & Trim(crucedoc) & "','"
   cad = cad & Trim(obliga) & "','"
   cad = cad & Trim(cuenta4) & "','"
   cad = cad & Trim(cuenta5) & "','"
   cad = cad & Trim(cuenta6) & "','"
   cad = cad & Trim(cuenta7) & "','"
   cad = cad & Trim(anticipo) & "','"
   cad = cad & Trim(cajachica) & "','"
   cad = cad & Trim(archivoe) & "','"
   cad = cad & Trim(local1) & "')"
   cn.Execute (cad)
   MsgBox "Adicion exitosa", 48, "Aviso"
   codigo.SetFocus
   Exit Sub
End If

If sw = 1 Then
   cad = "UPDATE tipo SET "
   cad = cad & "descripcio='" & Trim(descripcio) & "'"
   cad = cad & ",tipodoc = '" & Trim(tipodoc) & "'"
   cad = cad & ",sunat = '" & Trim(sunat) & "'"
   cad = cad & ",serie = '" & Trim(serie) & "'"
   cad = cad & ",numero = '" & Trim(numero) & "'"
   cad = cad & ",nrolineas = '" & Trim(nrolineas) & "'"
   cad = cad & ",puerto = '" & Trim(puerto) & "'"
   cad = cad & ",bodega = '" & Trim(bodega) & "'"
   cad = cad & ",contable = '" & Trim(contable) & "'"
   cad = cad & ",archivo = '" & Trim(archivo) & "'"
   cad = cad & ",grupo = '" & Trim(grupo) & "'"
   cad = cad & ",flage = '" & Trim(flage) & "'"
   cad = cad & ",te = '" & Trim(te) & "'"
   cad = cad & ",ts = '" & Trim(ts) & "'"
   cad = cad & ",cuenta1 = '" & Trim(cuenta1) & "'"
   cad = cad & ",cuenta2 = '" & Trim(cuenta2) & "'"
   cad = cad & ",cuenta3 = '" & Trim(cuenta3) & "'"
   cad = cad & ",crucedoc = '" & Trim(crucedoc) & "'"
   cad = cad & ",obliga = '" & Trim(obliga) & "'"
   cad = cad & ",cuenta4 = '" & Trim(cuenta4) & "'"
   cad = cad & ",cuenta5 = '" & Trim(cuenta5) & "'"
   cad = cad & ",cuenta6 = '" & Trim(cuenta6) & "'"
   cad = cad & ",cuenta7 = '" & Trim(cuenta7) & "'"
   cad = cad & ",anticipo = '" & Trim(anticipo) & "'"
   cad = cad & ",cajachica = '" & Trim(cajachica) & "'"
   cad = cad & ",archivoe = '" & Trim(archivoe) & "'"
   
   cad = cad & " WHERE   tipo='" & Trim(codigo) & "'"
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


rsexiste.Open "SELECT * FROM tipo where   tipo='" & Trim(codigo) & "'", cn, adOpenKeyset, adLockOptimistic
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
If Len(archivo) = 0 Then
   archivo.SetFocus
   Exit Function
End If
If Len(archivoe) = 0 Then
   archivoe = archivo
End If

valida = 1
End Function

Private Sub tipo_Change()

End Sub

Private Sub nrolineas_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
puerto.SetFocus

End Sub

Private Sub nrolineas_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   numero.SetFocus
   Exit Sub
End If

End Sub

Private Sub numero_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
nrolineas.SetFocus

End Sub

Private Sub numero_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   serie.SetFocus
   Exit Sub
End If

End Sub

Private Sub puerto_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
bodega.SetFocus

End Sub

Private Sub puerto_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   nrolineas.SetFocus
   Exit Sub
End If

End Sub

Private Sub serie_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
numero.SetFocus

End Sub

Private Sub serie_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   sunat.SetFocus
   Exit Sub
End If

End Sub

Private Sub sunat_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
serie.SetFocus

End Sub

Private Sub sunat_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   tipodoc.SetFocus
   Exit Sub
End If

End Sub

Private Sub tipodoc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
sunat.SetFocus

End Sub
