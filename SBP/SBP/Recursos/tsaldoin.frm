VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tsaldoin 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saldo Inicial"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   15075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   15075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Seleccione Un Periodo"
      Height          =   3495
      Left            =   3360
      TabIndex        =   76
      Top             =   1560
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton Command13 
         Caption         =   "Close"
         Height          =   495
         Left            =   2400
         TabIndex        =   79
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Acepta"
         Height          =   495
         Left            =   4080
         TabIndex        =   78
         Top             =   2640
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid dbgrid12 
         Height          =   2055
         Left            =   120
         TabIndex        =   77
         Top             =   360
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   3625
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   25
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
   Begin VB.TextBox bodega 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   75
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox local1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   74
      Top             =   0
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Procesando"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   1920
      TabIndex        =   73
      Top             =   1680
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Cargar productos"
      Height          =   4815
      Left            =   960
      TabIndex        =   62
      Top             =   960
      Visible         =   0   'False
      Width           =   9495
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Poner en Saldo Anterior"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2400
         TabIndex        =   71
         Top             =   1920
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Poner en Cantidad"
         Height          =   375
         Left            =   2160
         TabIndex        =   70
         Top             =   1320
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&P.Procesar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&X.Cerrar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Cargar Saldos Actuales"
         Height          =   255
         Left            =   2400
         TabIndex        =   64
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Costos"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2160
         TabIndex        =   63
         Top             =   2400
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro Prod.Cargados"
         Height          =   375
         Left            =   5400
         TabIndex        =   69
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label contador 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   68
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seleccione una opcion para incluir en la carga"
         Height          =   375
         Left            =   1080
         TabIndex        =   67
         Top             =   840
         Width           =   3615
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   "&CargaExcell"
      Enabled         =   0   'False
      Height          =   375
      Left            =   13320
      TabIndex        =   61
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Lista Precios"
      Height          =   3735
      Left            =   2040
      TabIndex        =   57
      Top             =   1680
      Visible         =   0   'False
      Width           =   8295
      Begin VB.CommandButton Command10 
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
         Left            =   7440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tsaldoin.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Borrar registro"
         Top             =   360
         Width           =   735
      End
      Begin MSDBGrid.DBGrid DBGrid4 
         Height          =   3255
         Left            =   120
         OleObjectBlob   =   "tsaldoin.frx":1212
         TabIndex        =   59
         Top             =   360
         Width           =   7215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ingreso de Lineas"
      Height          =   3255
      Left            =   2640
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton Command9 
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
         Left            =   5400
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tsaldoin.frx":2275
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Borrar registro"
         Top             =   2400
         Width           =   735
      End
      Begin VB.CommandButton Command8 
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
         Left            =   6240
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tsaldoin.frx":3487
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Grabar registro"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox t16 
         Height          =   375
         Left            =   6240
         MaxLength       =   2
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t15 
         Height          =   375
         Left            =   6240
         MaxLength       =   2
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t14 
         Height          =   375
         Left            =   6240
         MaxLength       =   2
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox t13 
         Height          =   375
         Left            =   6240
         MaxLength       =   2
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t12 
         Height          =   375
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t11 
         Height          =   375
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t10 
         Height          =   375
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox t9 
         Height          =   375
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t8 
         Height          =   375
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t7 
         Height          =   375
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t6 
         Height          =   375
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox t5 
         Height          =   375
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t4 
         Height          =   375
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t3 
         Height          =   375
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t2 
         Height          =   375
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox t1 
         Height          =   375
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.Label sumax 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   56
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label linea 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4440
         TabIndex        =   53
         Top             =   360
         Width           =   855
      End
      Begin VB.Label nt16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   52
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label nt15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   51
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   50
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label nt13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   49
         Top             =   840
         Width           =   735
      End
      Begin VB.Label nt12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   48
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label nt11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   47
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   46
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label nt9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   45
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label30 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   44
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label29 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   43
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label28 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   42
         Top             =   840
         Width           =   735
      End
      Begin VB.Label nt8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   41
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label nt7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   40
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   39
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label nt5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   38
         Top             =   840
         Width           =   735
      End
      Begin VB.Label nt4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   37
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label nt3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   36
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   35
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label nt1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   34
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tallas"
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   840
         Width           =   975
      End
      Begin VB.Label nlinea 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   32
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Linea"
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&EjecutaCondicion"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5640
      TabIndex        =   11
      Top             =   6720
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   10
      Text            =   "%"
      Top             =   6720
      Width           =   1935
   End
   Begin VB.ComboBox Combo3 
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
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   7080
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   6720
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Retorno"
      Enabled         =   0   'False
      Height          =   375
      Left            =   13320
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&CargarProductos"
      Enabled         =   0   'False
      Height          =   375
      Left            =   13320
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox fecha 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid dbGrid1 
      Height          =   5415
      Left            =   120
      TabIndex        =   72
      Top             =   1200
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   9551
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "Producto"
         Caption         =   "Producto"
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
      BeginProperty Column02 
         DataField       =   "Unidad"
         Caption         =   "Unidad"
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
      BeginProperty Column03 
         DataField       =   "Factor"
         Caption         =   "Factor"
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
      BeginProperty Column04 
         DataField       =   "Costo"
         Caption         =   "Costo"
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
      BeginProperty Column05 
         DataField       =   "Cantidad1"
         Caption         =   "Cantidad1"
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
      BeginProperty Column06 
         DataField       =   "Familia"
         Caption         =   "Familia"
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
      BeginProperty Column07 
         DataField       =   "Linea"
         Caption         =   "Linea"
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
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5729.953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   659.906
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   14280
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   60
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ojo. Si existe linea no puede ingresar cantidad,si no el contenido de la linea"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   120
      TabIndex        =   54
      Top             =   7560
      Width           =   7335
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ordenado"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Seleccionar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Almacen"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Periodo dd/mm/aaaa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
   Begin VB.Menu dlo3431 
      Caption         =   "&Opciones"
      Begin VB.Menu dbo834 
         Caption         =   "&3.Borrar la carga Actual Productos"
      End
      Begin VB.Menu con78343 
         Caption         =   "&4.Cargar Conteo Fisico Generado"
      End
   End
   Begin VB.Menu ldso23 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tsaldoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xproducto As String
Private Type campo_precio
    unidad As String
    factor As String
    precio As String
    costo As String
    margen As String
    stock As String
End Type
Dim tsnn1 As New ADODB.Recordset
Dim campo_precios(12) As campo_precio

Private Sub bodega_Click()
'Dim found As Integer
'found = busca_parame(extra_loquesea(bodega))
End Sub

Private Sub bodega_DblClick()
'Dim found As Integer
'found = busca_parame(extra_loquesea(bodega))
End Sub

Private Sub bodega_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim found As Integer
'found = busca_parame(extra_loquesea(bodega))
End Sub

Private Sub bodega_KeyPress(KeyAscii As Integer)
'Dim found As Integer
'found = busca_parame(extra_loquesea(bodega))
End Sub

Private Sub bodega_KeyUp(KeyCode As Integer, Shift As Integer)
'Dim found As Integer
'found = busca_parame(extra_loquesea(bodega))
End Sub

Private Sub Command1_Click()
Dim mytablex As New ADODB.Recordset
  If tsnn1.State = 1 Then tsnn1.Close
   tsnn1.Open "select *  from saldoini where local='" & local1 & "' and fecha='" & fecha & "' and bodega='" & bodega & "'", cn, adOpenStatic, adLockOptimistic
   If tsnn1.RecordCount = 0 Then
      'Exit Sub
   End If
   Set dbGrid1.DataSource = tsnn1
               habilita 0
               habilita1 1
               dbGrid1.SetFocus
               
   'Do
   'If tsnn1.EOF Then Exit Do
   
   '----------------------------------------------
   'Set mytablex = Nothing
   'mytablex.Open "select *  from producto where producto='" & tsnn1.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic
   'If mytablex.RecordCount > 0 Then
   '   tsnn1.Fields("descripcio") = "" & mytablex.Fields("descripcio")
   'End If
   'mytablex.Close
   '----------------------------------------------
   
   'tsnn1.MoveNext
   'Loop
   

End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub Command10_Click()
Frame5.Visible = False
dbGrid1.SetFocus

End Sub




Private Sub Command11_Click()
Dim pdtcarga As String
CommonDialog1.DialogTitle = "Seleccione un archivo Grafico"
CommonDialog1.InitDir = globaldir & "\excell"
CommonDialog1.Filter = "Archivos Excell|*.xls"
CommonDialog1.ShowOpen
'Si seleccionamos un archivo mostramos la ruta
If CommonDialog1.FileName <> "" Then
   pdtcarga = CommonDialog1.FileName
   Call Excel_a_Access(pdtcarga)
   Command1_Click
   'foto = LoadPicture(fotonombre)
Else
   'Si no mostramos un texto de advertencia de que no se seleccionó _   ninguno, ya que FileName devuelve una cadena vacía
   'Label1 = "No se seleccionó ningún archivo"
End If

End Sub

Private Sub Command12_Click()


Dim buf1 As String
Dim mytablex As New ADODB.Recordset
Dim mytabley As New ADODB.Recordset
Dim mytablez As New ADODB.Recordset
Dim buf As String
On Error GoTo cmdsel90
buf1 = "" & dbgrid12.columns("periodo")
mytablex.Open "select * from pdtde where periodo='" & buf1 & "' and local='" & local1 & "' and bodega='" & bodega & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount = 0 Then
   MsgBox "NO existen dato ", 48, "Aviso"
   mytablex.Close
   Exit Sub
End If
MsgBox "Nro Registro >" & mytablex.RecordCount
Do
If mytablex.EOF Then Exit Do
buf = "select * from saldoini where local='" & local1 & "' and fecha='" & fecha & "' and bodega='" & bodega & "' and producto='" & mytablex.Fields("producto") & "'"
           mytabley.Open buf, cn, adOpenStatic, adLockOptimistic
           If mytabley.RecordCount = 0 Then
              mytabley.AddNew
           End If
                      Set mytablez = Nothing
                      mytablez.Open "select costou from producto where producto='" & mytablex.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic
                      If mytablez.RecordCount > 0 Then
                         mytabley.Fields("costo") = Val("" & mytablez.Fields("costou"))
                      End If
                      mytablez.Close

              mytabley.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
              mytabley.Fields("bodega") = bodega
              mytabley.Fields("producto") = "" & mytablex.Fields("producto")
              mytabley.Fields("descripcio") = "" & mytablex.Fields("descripcio")
              mytabley.Fields("unidad") = "" & mytablex.Fields("unidad")
              mytabley.Fields("factor") = "" & mytablex.Fields("factor")
              mytabley.Fields("familia") = "" & mytablex.Fields("familia")
              mytabley.Fields("cantidad1") = Val("" & mytabley.Fields("cantidad1")) + Val("" & mytablex.Fields("saldoant"))
              mytabley.Fields("local") = local1
           mytabley.Update
           mytabley.Close
           Set mytabley = Nothing
mytablex.MoveNext
Loop
mytablex.Close

MsgBox "Proceso Terminado ", 48, "Aviso"
Frame3.Visible = False
Exit Sub
cmdsel90:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub



End Sub

Private Sub Command13_Click()
Frame3.Visible = False
End Sub

Private Sub Command2_Click()
Frame1.Visible = True


End Sub

Private Sub Command3_Click()
habilita 1
habilita1 0
Command1.SetFocus
End Sub

Private Sub Command4_Click()
Dim buf As String
buf = "select * from saldoini where local='" & local1 & "' and fecha='" & fecha & "' and bodega='" & bodega & "'"
If Combo2 <> "%" Then
buf = buf & " and " & Combo2 & " like '" & Text1 & "%'"
End If
If Combo3 <> "%" Then
   If Combo3 = "PRODUCTO" Then
      buf = buf & " order by producto"
   Else
   buf = buf & " order by " & Combo3
   End If
End If

Set tsnn1 = Nothing

If tsnn1.State = 1 Then
   Set tsnn1 = Nothing
   tsnn1.Close
End If
   tsnn1.Open buf, cn, adOpenStatic, adLockOptimistic
   If tsnn1.RecordCount = 0 Then
      Exit Sub
   End If
   Set dbGrid1.DataSource = tsnn1
   dbGrid1.refresh
   dbGrid1.SetFocus

End Sub

Private Sub Command5_Click()
Dim sw As Integer
Dim sdx As Double
Dim mytablex As New ADODB.Recordset
Dim mytabley As New ADODB.Recordset
Dim mytablez As New ADODB.Recordset
Dim xsaldo As Double
Dim vr
If MsgBox("Desea Realizar el proceso", 1, "Aviso") <> 1 Then Exit Sub

sdx = 0
mytablex.Open "select * from producto ", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount = 0 Then
   MsgBox "No existen productos ", 48, "Aviso"
   Exit Sub
End If
Command7.Visible = True

Do
If mytablex.EOF Then Exit Do
vr = DoEvents()
If Command7.Visible = False Then
   Exit Do
End If
sdx = sdx + 1
contador = Format(sdx, "000000")
xsaldo = 0
mytablez.Open "select * from almacen where local='" & local1 & "' and producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & bodega & "'", cn, adOpenStatic, adLockOptimistic
If mytablez.RecordCount > 0 Then
   xsaldo = Val("" & mytablez.Fields("saldo"))
End If
mytablez.Close
If mytabley.State = 1 Then
   mytabley.Close
   Set mytabley = Nothing
End If
mytabley.Open "select * from saldoini where local='" & local1 & "' and producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & bodega & "' and fecha='" & fecha & "'", cn, adOpenStatic, adLockOptimistic
If mytabley.RecordCount = 0 Then
   mytabley.AddNew
   mytabley.Fields("producto") = "" & mytablex.Fields("producto")
   mytabley.Fields("descripcio") = Mid$("" & mytablex.Fields("descripcio"), 1, 60)
   mytabley.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
   mytabley.Fields("bodega") = Trim("" & bodega)
   mytabley.Fields("local") = Trim(local1)
   mytabley.Fields("unidad") = "" & mytablex.Fields("unidad")
   mytabley.Fields("factor") = "" & mytablex.Fields("factor")
   mytabley.Fields("familia") = "" & mytablex.Fields("familia")
   mytabley.Fields("linea") = "" & mytablex.Fields("linea")
   mytabley.Fields("local") = Trim(local1)
   mytabley.Fields("costo") = Val("" & mytablex.Fields("costou"))
   If Check1.Value = 1 Then
      If Option1.Value = True Then
         mytabley.Fields("cantidad") = xsaldo
      End If
      If Option2.Value = True Then
         mytabley.Fields("saldoant") = xsaldo
      End If
   End If
   If Check1.Value = 0 Then
      mytabley.Fields("cantidad") = 0
   End If
   'mytabley.Fields("cantidad") = 0
   mytabley.Update
Else
   'mytabley.Edit
   mytabley.Fields("producto") = "" & mytablex.Fields("producto")
   mytabley.Fields("descripcio") = "" & mytablex.Fields("descripcio")
   mytabley.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
   mytabley.Fields("bodega") = Trim(bodega)
   mytabley.Fields("unidad") = "" & mytablex.Fields("unidad")
   mytabley.Fields("factor") = "" & mytablex.Fields("factor")
   mytabley.Fields("familia") = "" & mytablex.Fields("familia")
   mytabley.Fields("linea") = "" & mytablex.Fields("linea")
   mytabley.Fields("costo") = Val("" & mytablex.Fields("costou"))
   mytabley.Fields("local") = Trim(local1)
   If Check1.Value = 1 Then
      If Option1.Value = True Then
         mytabley.Fields("cantidad") = xsaldo
      End If
   '   If Option2.Value = True Then
   '      mytabley.Fields("saldoant") = xsaldo
   '   End If
   End If
   If Check1.Value = 0 Then
      mytabley.Fields("cantidad") = 0
   End If
   'mytabley.Fields("cantidad") = 0
   mytabley.Update
End If
mytabley.Close
mytablex.MoveNext
Loop
Command7.Visible = False
'mytablez.Close
'mytabley.Close
mytablex.Close

Command6_Click
Command1_Click
dbGrid1.SetFocus
End Sub

Private Sub Command6_Click()
Frame1.Visible = False
End Sub

Private Sub Command7_Click()
Command7.Visible = False

End Sub

Private Sub Command8_Click()
Dim sdx As Double
suma_xx

dbGrid1.columns(8) = Val(t1)
dbGrid1.columns(9) = Val(t2)
dbGrid1.columns(10) = Val(t3)
dbGrid1.columns(11) = Val(t4)
dbGrid1.columns(12) = Val(t5)
dbGrid1.columns(13) = Val(t6)
dbGrid1.columns(14) = Val(t7)
dbGrid1.columns(15) = Val(t8)
dbGrid1.columns(16) = Val(t9)
dbGrid1.columns(17) = Val(t10)
dbGrid1.columns(18) = Val(t11)
dbGrid1.columns(19) = Val(t12)
dbGrid1.columns(20) = Val(t13)
dbGrid1.columns(21) = Val(t14)
dbGrid1.columns(22) = Val(t15)
dbGrid1.columns(23) = Val(t16)
sdx = Val(t1) + Val(t2) + Val(t3) + Val(t4) + Val(t5) + Val(t6) + Val(t7) + Val(t8) + Val(t9) + Val(t10) + Val(t11) + Val(t12) + Val(t13) + Val(t14) + Val(t15) + Val(t16)
dbGrid1.columns(4) = sdx
Command9_Click

End Sub

Private Sub Command9_Click()
Frame2.Visible = False
dbGrid1.SetFocus
End Sub

Private Sub con78343_Click()
Dim mytablex As New ADODB.Recordset
   Frame3.Visible = True
   mytablex.Open "select *  from periodo", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount = 0 Then
      Exit Sub
   End If
   Set dbgrid12.DataSource = mytablex
   dbgrid12.SetFocus

End Sub

Private Sub dbgrid1_AfterColUpdate(ByVal ColIndex As Integer)
Select Case ColIndex
       Case 5
             tsnn1.Update
       'MsgBox "" & dbGrid1.columns(4)
End Select

End Sub

Private Sub dbgrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If ColIndex <> 4 And ColIndex <> 5 Then
   Cancel = True
   Exit Sub
End If
Select Case ColIndex

           
       Case 4, 5
            'MsgBox "" & DBGrid1.columns(4)
            If Len("" & dbGrid1.columns(7)) > 0 Then
               Cancel = True
               Exit Sub
            End If
            
End Select
End Sub

Private Sub dbgrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Select Case ColIndex
       Case 5
       'MsgBox OldValue
End Select
End Sub

Private Sub dbgrid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   Exit Sub
   If Len(dbGrid1.columns(0)) > 0 And dbGrid1.Col = 2 Then
      xproducto = "" & dbGrid1.columns(0)
      carga_dbgrid4
   End If
   Exit Sub
End If
If KeyCode = &H71 Then  'f2
   If Len(dbGrid1.columns(7)) > 0 And Len(dbGrid1.columns(0)) > 0 Then
      ingreso_tallas "" & dbGrid1.columns(7)
   End If
End If


End Sub

Private Sub DBGrid4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Command10_Click
   Exit Sub
End If
If KeyCode = 13 Then
   If Len("" & dbgrid4.columns(0)) > 0 And Val("" & dbgrid4.columns(1)) > 0 And Len("" & dbgrid4.columns(3)) > 0 Then
      'Data1.Recordset.Edit
      dbGrid1.columns("und") = "" & dbgrid4.columns(0)
      dbGrid1.columns("factor") = "" & dbgrid4.columns(1)
      'Data1.Recordset.Fields("costo") = "" & DBGrid4.Columns(3)
      'dbGrid1.Columns.Update
      Command10_Click
   End If
End If

End Sub

Private Sub DBGrid4_UnboundReadData(ByVal RowBuf As MSDBGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim dR As Integer
Dim row_num As Integer
Dim R As Integer
Dim rows_returned As Integer
If ReadPriorRows Then
        dR = -1
    Else
        dR = 1
    End If
    If IsNull(StartLocation) Then
        If ReadPriorRows Then
           row_num = RowBuf.RowCount - 1
           'row_num = 9
        Else
           row_num = 0
        End If
    Else
        row_num = CLng(StartLocation) + dR
    End If
    rows_returned = 0
    For R = 0 To RowBuf.RowCount - 1
        If row_num < 0 Or row_num > 9 Then Exit For
        RowBuf.Value(R, 0) = campo_precios(row_num).unidad
        RowBuf.Value(R, 1) = campo_precios(row_num).factor
        RowBuf.Value(R, 2) = campo_precios(row_num).precio
        RowBuf.Value(R, 3) = campo_precios(row_num).costo
        RowBuf.Value(R, 4) = campo_precios(row_num).margen
        RowBuf.Value(R, 5) = campo_precios(row_num).stock
        RowBuf.Bookmark(R) = row_num
        row_num = row_num + dR
        rows_returned = rows_returned + 1
   Next R
   RowBuf.RowCount = rows_returned

End Sub

Private Sub dbo834_Click()

On Error GoTo cmd23_err
If Frame1.Visible = True Then Exit Sub
'If Command1.Enabled = True Then Exit Sub
If MsgBox("Se borrara la data de :" & fecha, 1, "Aviso") <> 1 Then Exit Sub
cn.Execute "DELETE FROM saldoini where local='" & local1 & "' and fecha='" & fecha & "' and bodega='" & bodega & "'"
MsgBox "Proceso Realizado ", 48, "Aviso"
Command1_Click
Exit Sub
cmd23_err:
Exit Sub
End Sub

Private Sub df883_Click()

End Sub

Private Sub fecha_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If Len(fecha) = 0 Then
   carga_fecha
End If
If KeyAscii = 27 Then
   Exit Sub
End If
If Len(fecha) <> 10 Then Exit Sub
If Not IsDate(fecha) Then Exit Sub

End Sub


Private Sub fkli3e3_Click()
If Frame1.Visible = False Then Exit Sub
If Command1.Enabled = True Then Exit Sub
trecalcu.Show 1
End Sub

Private Sub Form_Activate()
'found = busca_parame(bodega)
Command1_Click

End Sub

Private Sub Form_Load()
Combo2.Clear
Combo2.AddItem "%"
Combo2.AddItem "DESCRIPCIO"
Combo2.AddItem "PRODUCTO"
Combo2.AddItem "FAMILIA"
Combo2.AddItem "LINEA"
Combo2.ListIndex = 0

Combo3.Clear
Combo3.AddItem "%"
Combo3.AddItem "DESCRIPCIO"
Combo3.AddItem "PRODUCTO"
Combo3.AddItem "FAMILIA"
Combo3.AddItem "LINEA"
Combo3.ListIndex = 0

End Sub

Private Sub ldso23_Click()
If Command7.Visible = True Then
   Command7.Visible = False
   Exit Sub
End If

If Frame1.Visible = True Then
   Frame1.Visible = False
   Exit Sub
End If

'If Command1.Enabled = False Then
'   Command1.Enabled = True
'   habilita 1
'   habilita1 0
'   Command1.SetFocus
'   Exit Sub
'End If

tsaldoin.Hide
Unload tsaldoin
End Sub
Sub carga_fecha()
Dim mytablex As New ADODB.Recordset
mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   fecha = "" & mytablex.Fields("saldoini")
End If
mytablex.Close
End Sub
Sub habilita(sw As Integer)
Dim xsw
If sw = 0 Then
   xsw = True
End If
If sw = 1 Then
   xsw = False
End If
Command11.Enabled = xsw
Command2.Enabled = xsw
Command3.Enabled = xsw
Command4.Enabled = xsw
dbGrid1.Enabled = xsw
End Sub
Sub habilita1(sw As Integer)
Dim xsw
If sw = 0 Then
   xsw = True
End If
If sw = 1 Then
   xsw = False
End If
'Command1.Enabled = xsw
local1.Enabled = xsw
'fecha.Enabled = xsw
bodega.Enabled = xsw

End Sub
Sub ingreso_tallas(buf As String)
Dim found As Integer
linea = buf
found = busca_linea(buf)
If found = 0 Then Exit Sub
pone_tallas
Frame2.Visible = True
t1.SetFocus
End Sub
Function busca_linea(buf As String)
Dim mytablex As New ADODB.Recordset
mytablex.Open "select * from linea where linea='" & buf & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   busca_linea = 1
   nt1 = "" & mytablex.Fields("t1")
   nt2 = "" & mytablex.Fields("t2")
   nt3 = "" & mytablex.Fields("t3")
   nt4 = "" & mytablex.Fields("t4")
   nt5 = "" & mytablex.Fields("t5")
   nt6 = "" & mytablex.Fields("t6")
   nt7 = "" & mytablex.Fields("t7")
   nt8 = "" & mytablex.Fields("t8")
   nt9 = "" & mytablex.Fields("t9")
   nt10 = "" & mytablex.Fields("t10")
   nt11 = "" & mytablex.Fields("t11")
   nt12 = "" & mytablex.Fields("t12")
   nt13 = "" & mytablex.Fields("t13")
   nt14 = "" & mytablex.Fields("t14")
   nt15 = "" & mytablex.Fields("t15")
   nt16 = "" & mytablex.Fields("t16")
End If
mytablex.Close
 
End Function

Private Sub local1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
bodega.SetFocus
End Sub

Private Sub pero83453_Click()
If Frame1.Visible = True Then Exit Sub
End Sub


Private Sub t1_Change()
suma_xx
End Sub
Sub suma_xx()
Dim sdx As Double
sdx = Val(t1) + Val(t2) + Val(t3) + Val(t4) + Val(t5) + Val(t6) + Val(t7) + Val(t8) + Val(t9) + Val(t10) + Val(t11) + Val(t12) + Val(t13) + Val(t14) + Val(t15) + Val(t16)
sumax = Format(sdx, "0")
End Sub

Private Sub t1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
suma_xx
t2.SetFocus
End Sub

Private Sub t10_Change()
suma_xx
End Sub

Private Sub t10_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
suma_xx
t11.SetFocus

End Sub

Private Sub t10_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   t9.SetFocus
   Exit Sub
End If

End Sub

Private Sub t11_Change()
suma_xx
End Sub

Private Sub t11_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
suma_xx
t12.SetFocus

End Sub

Private Sub t11_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   t10.SetFocus
   Exit Sub
End If

End Sub

Private Sub t12_Change()
suma_xx
End Sub

Private Sub t12_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
suma_xx
t13.SetFocus

End Sub

Private Sub t12_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   t11.SetFocus
   Exit Sub
End If

End Sub

Private Sub t13_Change()
suma_xx
End Sub

Private Sub t13_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
suma_xx
t14.SetFocus

End Sub

Private Sub t13_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   t12.SetFocus
   Exit Sub
End If

End Sub

Private Sub t14_Change()
suma_xx
End Sub

Private Sub t14_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
suma_xx
t15.SetFocus

End Sub

Private Sub t14_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   t13.SetFocus
   Exit Sub
End If

End Sub

Private Sub t15_Change()
suma_xx
End Sub

Private Sub t15_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
suma_xx
t16.SetFocus

End Sub

Private Sub t15_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   t14.SetFocus
   Exit Sub
End If

End Sub

Private Sub t16_Change()
suma_xx
End Sub

Private Sub t16_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
suma_xx
Command8_Click

End Sub

Private Sub t16_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   t15.SetFocus
   Exit Sub
End If

End Sub

Private Sub t2_Change()
suma_xx
End Sub

Private Sub t2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
suma_xx
t3.SetFocus

End Sub

Private Sub t2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   t1.SetFocus
   Exit Sub
End If

End Sub

Private Sub t3_Change()
suma_xx
End Sub

Private Sub t3_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
suma_xx
t4.SetFocus

End Sub

Private Sub t3_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   t2.SetFocus
   Exit Sub
End If

End Sub

Private Sub t4_Change()
suma_xx
End Sub

Private Sub t4_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
suma_xx
t5.SetFocus

End Sub

Private Sub t4_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   t3.SetFocus
   Exit Sub
End If

End Sub

Private Sub t5_Change()
suma_xx
End Sub

Private Sub t5_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
suma_xx
t6.SetFocus

End Sub

Private Sub t5_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   t4.SetFocus
   Exit Sub
End If

End Sub

Private Sub t6_Change()
suma_xx
End Sub

Private Sub t6_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
suma_xx
t7.SetFocus

End Sub

Private Sub t6_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   t5.SetFocus
   Exit Sub
End If

End Sub

Private Sub t7_Change()
suma_xx
End Sub

Private Sub t7_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
suma_xx
t8.SetFocus

End Sub

Private Sub t7_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   t6.SetFocus
   Exit Sub
End If

End Sub

Private Sub t8_Change()
suma_xx
End Sub

Private Sub t8_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
suma_xx
t9.SetFocus

End Sub

Private Sub t8_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   t7.SetFocus
   Exit Sub
End If

End Sub

Private Sub t9_Change()
suma_xx
End Sub
Sub pone_tallas()
t1 = "" & tsnn1.Fields("t1")
t2 = "" & tsnn1.Fields("t2")
t3 = "" & tsnn1.Fields("t3")
t4 = "" & tsnn1.Fields("t4")
t5 = "" & tsnn1.Fields("t5")
t6 = "" & tsnn1.Fields("t6")
t7 = "" & tsnn1.Fields("t7")
t8 = "" & tsnn1.Fields("t8")
t9 = "" & tsnn1.Fields("t9")
t10 = "" & tsnn1.Fields("t10")
t11 = "" & tsnn1.Fields("t11")
t12 = "" & tsnn1.Fields("t12")
t13 = "" & tsnn1.Fields("t13")
t14 = "" & tsnn1.Fields("t14")
t15 = "" & tsnn1.Fields("t15")
t16 = "" & tsnn1.Fields("t16")
End Sub

Sub carga_dbgrid4()
Dim i As Integer

Dim mytablex As New ADODB.Recordset
Dim mytabley As New ADODB.Recordset
Dim sw As Integer
Dim xbodega As String
Dim xsaldo As Double
Dim xbuf As String
Dim xcosto As Double
Dim xcostou As Double
Dim xfactor As Double
Dim xunidad As String
Dim xmargen As Double
On Error GoTo cmd89012_err
For i = 0 To 9
    campo_precios(i).unidad = ""
    campo_precios(i).factor = ""
    campo_precios(i).precio = ""
    campo_precios(i).costo = ""
    campo_precios(i).margen = ""
    campo_precios(i).stock = ""
Next i
xcostou = 0
xunidad = "UND"
xfactor = 1
xbodega = bodega
xsaldo = 0
xcosto = 0
sw = 0

   mytabley.Open "SELECT * FROM almacen where  local='" & local1 & "' and producto='" & xproducto & "' and bodega='" & xbodega & "'", cn, adOpenKeyset, adLockOptimistic
   If mytabley.RecordCount > 0 Then  'si existe
      xsaldo = Val("" & mytabley.Fields("saldo"))
   End If
   mytabley.Close
   
   mytablex.Open "SELECT * FROM producto where  producto='" & xproducto & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount > 0 Then  'si existe
      xcostou = Val("" & mytablex.Fields("costou"))
      xfactor = Val("" & mytablex.Fields("factor"))
      xunidad = "" & mytablex.Fields("unidad")
   End If
   mytablex.Close
   
   mytablex.Open "SELECT * FROM precios where  producto='" & xproducto & "' and local='" & local1 & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount > 0 Then  'si existe
      'MsgBox "Hola"
      xcosto = xcostou
      campo_precios(0).unidad = xunidad
      campo_precios(0).factor = xfactor
      campo_precios(0).precio = "" '& mytablex.Fields("costou")
      campo_precios(0).costo = xcostou
      xbuf = calcula_saldo(xsaldo, xfactor)
      campo_precios(0).stock = "" & xbuf
      xmargen = 0
      campo_precios(0).margen = "" & xmargen
      '----------------------------------------------
      xcosto = 0
      If Val("" & mytablex.Fields("factor1")) > 0 Then
         xcosto = xcostou / xfactor
         xcosto = xcosto * Val("" & mytablex.Fields("factor1"))
      campo_precios(1).unidad = "" & mytablex.Fields("unidad1")
      campo_precios(1).factor = "" & mytablex.Fields("factor1")
      campo_precios(1).precio = "" & mytablex.Fields("pventa1")
      campo_precios(1).costo = "" & xcosto
      xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor1")))
      campo_precios(1).stock = "" & xbuf
      xmargen = 0
      If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa1")) - xcosto) * 100) / xcosto
      End If
      campo_precios(1).margen = "" & xmargen
   '--------
   End If
   '---------
   If Val("" & mytablex.Fields("factor2")) > 0 Then
   campo_precios(2).unidad = "" & mytablex.Fields("unidad2")
   campo_precios(2).factor = "" & mytablex.Fields("factor2")
   campo_precios(2).precio = "" & mytablex.Fields("pventa2")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor2")))
   campo_precios(2).stock = "" & xbuf
   xcosto = 0
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor2"))
   campo_precios(2).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
      xmargen = ((Val("" & mytablex.Fields("pventa2")) - xcosto) * 100) / xcosto
   End If
   campo_precios(2).margen = "" & xmargen
   End If
   
   If Val("" & mytablex.Fields("factor3")) > 0 Then
   campo_precios(3).unidad = "" & mytablex.Fields("unidad3")
   campo_precios(3).factor = "" & mytablex.Fields("factor3")
   campo_precios(3).precio = "" & mytablex.Fields("pventa3")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor3")))
   campo_precios(3).stock = "" & xbuf
   xcosto = 0
   
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor3"))
   
   campo_precios(3).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa3")) - xcosto) * 100) / xcosto
         campo_precios(3).margen = "" & xmargen
   End If
   campo_precios(3).margen = "" & xmargen
   End If
   If Val("" & mytablex.Fields("factor4")) > 0 Then
   campo_precios(4).unidad = "" & mytablex.Fields("unidad4")
   campo_precios(4).factor = "" & mytablex.Fields("factor4")
   campo_precios(4).precio = "" & mytablex.Fields("pventa4")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor4")))
   campo_precios(4).stock = "" & xbuf
   xcosto = 0
   
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor4"))
   
   campo_precios(4).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa4")) - xcosto) * 100) / xcosto
   End If
   campo_precios(4).margen = "" & xmargen
   End If
   
   If Val("" & mytablex.Fields("factor5")) > 0 Then
   campo_precios(5).unidad = "" & mytablex.Fields("unidad5")
   campo_precios(5).factor = "" & mytablex.Fields("factor5")
   campo_precios(5).precio = "" & mytablex.Fields("pventa5")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor5")))
   campo_precios(5).stock = "" & xbuf
   xcosto = 0
   
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor5"))
   
   campo_precios(5).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa5")) - xcosto) * 100) / xcosto
   End If
   campo_precios(5).margen = "" & xmargen
   End If
   
   If Val("" & mytablex.Fields("factor6")) > 0 Then
   campo_precios(6).unidad = "" & mytablex.Fields("unidad6")
   campo_precios(6).factor = "" & mytablex.Fields("factor6")
   campo_precios(6).precio = "" & mytablex.Fields("pventa6")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor6")))
   campo_precios(6).stock = "" & xbuf
   xcosto = 0
   
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor6"))
   
   campo_precios(6).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa5")) - xcosto) * 100) / xcosto
         
   End If
   campo_precios(6).margen = "" & xmargen
   End If
   
   If Val("" & mytablex.Fields("factor7")) > 0 Then
   campo_precios(7).unidad = "" & mytablex.Fields("unidad7")
   campo_precios(7).factor = "" & mytablex.Fields("factor7")
   campo_precios(7).precio = "" & mytablex.Fields("pventa7")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor7")))
   campo_precios(7).stock = "" & xbuf
   xcosto = 0
   
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor7"))
   
   campo_precios(7).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa7")) - xcosto) * 100) / xcosto
   End If
   campo_precios(7).margen = "" & xmargen
   End If
   
   
   If Val("" & mytablex.Fields("factor8")) > 0 Then
   campo_precios(8).unidad = "" & mytablex.Fields("unidad8")
   campo_precios(8).factor = "" & mytablex.Fields("factor8")
   campo_precios(8).precio = "" & mytablex.Fields("pventa8")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor8")))
   campo_precios(8).stock = "" & xbuf
   xcosto = 0
   
      xcosto = xcostou / xfactor
      xcosto = xcosto * Val("" & mytablex.Fields("factor8"))
   
   campo_precios(8).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
      xmargen = ((Val("" & mytablex.Fields("pventa8")) - xcosto) * 100) / xcosto
   End If
   campo_precios(8).margen = "" & xmargen
   End If
   
   If Val("" & mytablex.Fields("factor9")) > 0 Then
   campo_precios(9).unidad = "" & mytablex.Fields("unidad9")
   campo_precios(9).factor = "" & mytablex.Fields("factor9")
   campo_precios(9).precio = "" & mytablex.Fields("pventa9")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor9")))
   campo_precios(9).stock = "" & xbuf
   xcosto = 0
   
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor9"))
   
   campo_precios(9).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa9")) - xcosto) * 100) / xcosto
         
   End If
   campo_precios(9).margen = "" & xmargen
   End If
   
   If Val("" & mytablex.Fields("factor10")) > 0 Then
   campo_precios(10).unidad = "" & mytablex.Fields("unidad10")
   campo_precios(10).factor = "" & mytablex.Fields("factor10")
   campo_precios(10).precio = "" & mytablex.Fields("pventa10")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor10")))
   campo_precios(10).stock = "" & xbuf
   xcosto = 0
   
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor10"))
   
   campo_precios(10).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa10")) - xcosto) * 100) / xcosto
   End If
   campo_precios(10).margen = "" & xmargen
   End If
   'margenes
   sw = 1
   
 End If
mytablex.Close

dbgrid4.refresh
Frame5.Visible = True
dbgrid4.SetFocus
Exit Sub
cmd89012_err:
MsgBox "Error en carga Grid " + error$, 48, "Aviso"
Exit Sub

End Sub


Private Sub t9_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
suma_xx
t10.SetFocus

End Sub

Private Sub t9_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   t8.SetFocus
   Exit Sub
End If

End Sub
Function busca_parame(buf As String)
Dim sdx As Double
Dim mytablex As New ADODB.Recordset
fecha = ""
mytablex.Open "select * from bodega where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   fecha = "" & mytablex.Fields("fecha")
   busca_parame = 1
End If
mytablex.Close
End Function
Private Sub Excel_a_Access(Path_XLS As String)


Dim Obj_Excel As Object
Dim Obj_Hoja As Object
Dim Fila_Actual As Double
Dim buf As String
Dim DATO As Variant

Dim mytablex As New ADODB.Recordset

    Screen.MousePointer = vbHourglass
    Set Obj_Excel = CreateObject("Excel.Application")
    Obj_Excel.Workbooks.Open FileName:=Path_XLS
    If Val(Obj_Excel.Application.Version) >= 8 Then
        Set Obj_Hoja = Obj_Excel.ActiveSheet
    Else
        Set Obj_Hoja = Obj_Excel
    End If
    
    
    For Fila_Actual = 1 To 10000
           If Len(Trim$(Obj_Hoja.Cells(Fila_Actual, 1))) = 0 Then
              Exit For
           End If
           buf = "select * from saldoini where local='" & local1 & "' and fecha='" & fecha & "' and bodega='" & bodega & "' and producto='" & Trim$(Obj_Hoja.Cells(Fila_Actual, 1)) & "'"
           mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
           If mytablex.RecordCount > 0 Then
              DATO = Trim$(Obj_Hoja.Cells(Fila_Actual, 2))
              mytablex.Fields("cantidad") = Val(DATO)
              mytablex.Update
           End If
           mytablex.Close
        
    Next
    
    Call Descargar_Objetos(Obj_Excel, Obj_Hoja)
    Screen.MousePointer = vbDefault
    MsgBox " Datos copiados ", vbInformation
Exit Sub
'Error
ErrSub:

Call Descargar_Objetos(Obj_Excel, Obj_Hoja)
MsgBox Err.Description, vbCritical
Screen.MousePointer = vbDefault
    
End Sub

'Descarga los objetos y los cierra
Sub Descargar_Objetos(Obj_Excel As Object, Obj_Hoja As Object)


    
    Obj_Excel.ActiveWorkbook.Close False
    Obj_Excel.Quit
    Set Obj_Hoja = Nothing
    Set Obj_Excel = Nothing

End Sub


