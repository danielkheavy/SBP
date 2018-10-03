VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tconteoo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Regularizaciones de Inventario"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   15525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   15525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consultas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   1275
      TabIndex        =   97
      Top             =   5040
      Visible         =   0   'False
      Width           =   13725
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
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   360
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command8 
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
         Left            =   8160
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   360
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   6735
         Left            =   120
         TabIndex        =   101
         Top             =   960
         Width           =   13320
         _ExtentX        =   23495
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
      Begin VB.Label Label56 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   102
         Top             =   7800
         Width           =   14175
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Height          =   8415
      Left            =   1155
      TabIndex        =   42
      Top             =   -225
      Visible         =   0   'False
      Width           =   13695
      Begin VB.TextBox cantidad1 
         Height          =   375
         Left            =   4080
         MaxLength       =   15
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox producto 
         Height          =   375
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox conteo 
         Height          =   375
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   2160
         Width           =   2055
      End
      Begin VB.CommandButton Command7 
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
         Height          =   855
         Left            =   9000
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tconteo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Borrar registro"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
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
         Height          =   975
         Left            =   9000
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tconteo.frx":1212
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Grabar registro"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox t1 
         Height          =   375
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox t2 
         Height          =   375
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox t3 
         Height          =   375
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox t4 
         Height          =   375
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox t5 
         Height          =   375
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox t6 
         Height          =   375
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox t7 
         Height          =   375
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox t8 
         Height          =   375
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox t9 
         Height          =   375
         Left            =   4680
         MaxLength       =   2
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox t10 
         Height          =   375
         Left            =   4680
         MaxLength       =   2
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox t11 
         Height          =   375
         Left            =   4680
         MaxLength       =   2
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox t12 
         Height          =   375
         Left            =   4680
         MaxLength       =   2
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox t13 
         Height          =   375
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox t14 
         Height          =   375
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox t15 
         Height          =   375
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox t16 
         Height          =   375
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto"
         Height          =   375
         Left            =   120
         TabIndex        =   96
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcio"
         Height          =   375
         Left            =   120
         TabIndex        =   95
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label descripcio 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2040
         TabIndex        =   94
         Top             =   600
         Width           =   6135
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unidad"
         Height          =   375
         Left            =   120
         TabIndex        =   93
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Conteo"
         Height          =   375
         Left            =   120
         TabIndex        =   92
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saldo"
         Height          =   375
         Left            =   120
         TabIndex        =   91
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label saldo 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   90
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Linea"
         Height          =   375
         Left            =   120
         TabIndex        =   89
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label linea 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1800
         TabIndex        =   88
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label unidad 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2040
         TabIndex        =   87
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label factor 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2040
         TabIndex        =   86
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label27 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "factor"
         Height          =   375
         Left            =   120
         TabIndex        =   85
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Limpiar"
         Height          =   375
         Left            =   4080
         TabIndex        =   84
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tallas"
         Height          =   375
         Left            =   120
         TabIndex        =   83
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label nt1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1080
         TabIndex        =   82
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label nt2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1080
         TabIndex        =   81
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label nt3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1080
         TabIndex        =   80
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label nt4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1080
         TabIndex        =   79
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label nt5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2520
         TabIndex        =   78
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label nt6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2520
         TabIndex        =   77
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label nt7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2520
         TabIndex        =   76
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label nt8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2520
         TabIndex        =   75
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label nt9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3960
         TabIndex        =   74
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label nt10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3960
         TabIndex        =   73
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label nt11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3960
         TabIndex        =   72
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label nt12 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3960
         TabIndex        =   71
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label nt13 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5400
         TabIndex        =   70
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label nt14 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5400
         TabIndex        =   69
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label nt15 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5400
         TabIndex        =   68
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label nt16 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5400
         TabIndex        =   67
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label nlinea 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3360
         TabIndex        =   66
         Top             =   2760
         Width           =   3495
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
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
         Height          =   495
         Left            =   120
         TabIndex        =   65
         Top             =   4680
         Width           =   975
      End
      Begin VB.Label sumax 
         BackColor       =   &H00E0E0E0&
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
         Height          =   495
         Left            =   1080
         TabIndex        =   64
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ojo. Si existe linea no puede ingresar cantidad,si no el contenido de la linea  . Unidades solamente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   63
         Top             =   7440
         Width           =   11175
      End
   End
   Begin VB.CommandButton Command13 
      Caption         =   "+"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13200
      TabIndex        =   41
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Cargar Productos x Grupos"
      Height          =   7455
      Left            =   17160
      TabIndex        =   32
      Top             =   7920
      Visible         =   0   'False
      Width           =   14655
      Begin VB.CommandButton label20 
         Height          =   2415
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Visible         =   0   'False
         Width           =   14415
      End
      Begin VB.ComboBox marca 
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
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   720
         Width           =   4935
      End
      Begin VB.ComboBox familia 
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
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SALIDA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   12240
         TabIndex        =   38
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   12240
         TabIndex        =   37
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Marca"
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
         TabIndex        =   36
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Familia"
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
         TabIndex        =   34
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CargaTablaProd"
      Height          =   375
      Left            =   13200
      TabIndex        =   31
      Top             =   3120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox vendedor 
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox numero 
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
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Sumar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   13200
      TabIndex        =   23
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   13320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6360
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data5 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   13320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Importacion desde Excell"
      Height          =   4695
      Left            =   17160
      TabIndex        =   16
      Top             =   4080
      Visible         =   0   'False
      Width           =   9375
      Begin VB.CommandButton Command12 
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
         Left            =   7560
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label ncount 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   90
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Cargar&DesdePdt"
      Enabled         =   0   'False
      Height          =   375
      Left            =   13200
      TabIndex        =   15
      Top             =   2760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&EjecutaCondicion"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Top             =   7800
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
      TabIndex        =   12
      Text            =   "%"
      Top             =   7800
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
      TabIndex        =   11
      Top             =   8160
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
      TabIndex        =   9
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Regresar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   13200
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Ir a Detalle"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.ComboBox bodega 
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox fecha 
      BackColor       =   &H00C0FFFF&
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
      Top             =   1200
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid dbGrid2 
      Height          =   5655
      Left            =   120
      TabIndex        =   40
      Top             =   2040
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   9975
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
      ColumnCount     =   16
      BeginProperty Column00 
         DataField       =   "producto"
         Caption         =   "Producto"
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
         Caption         =   "Und"
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
         DataField       =   "factor"
         Caption         =   "Fac"
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
         DataField       =   "Cantidad"
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
      BeginProperty Column05 
         DataField       =   "saldoant"
         Caption         =   "Conteo"
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
      BeginProperty Column07 
         DataField       =   "Faltante"
         Caption         =   "Faltante"
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
      BeginProperty Column08 
         DataField       =   "Sobrante"
         Caption         =   "Sobrante"
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
      BeginProperty Column09 
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
      BeginProperty Column10 
         DataField       =   "Local"
         Caption         =   "Local"
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
      BeginProperty Column11 
         DataField       =   "fecha"
         Caption         =   "Fecha"
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
         DataField       =   "Bodega"
         Caption         =   "Bodega"
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
      BeginProperty Column13 
         DataField       =   "Numero"
         Caption         =   "Numero"
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
         DataField       =   "vendedor"
         Caption         =   "Vendedor"
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4635.213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   555.024
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   884.976
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1094.74
         EndProperty
      EndProperty
   End
   Begin VB.Label LOCAL1 
      BackColor       =   &H80000002&
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
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2280
      TabIndex        =   103
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label modelo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4320
      TabIndex        =   30
      Top             =   120
      Width           =   105
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Responsable"
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
      TabIndex        =   29
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Numero"
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
      TabIndex        =   28
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Conteo"
      Height          =   375
      Left            =   8160
      TabIndex        =   27
      Top             =   8160
      Width           =   975
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo"
      Height          =   375
      Left            =   8160
      TabIndex        =   26
      Top             =   7800
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sobrante"
      Height          =   375
      Left            =   10680
      TabIndex        =   25
      Top             =   8160
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Faltante"
      Height          =   375
      Left            =   10680
      TabIndex        =   24
      Top             =   7800
      Width           =   975
   End
   Begin VB.Label wsobrante 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11640
      TabIndex        =   22
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label wfaltante 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11640
      TabIndex        =   21
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label wconteo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   9120
      TabIndex        =   20
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label wsaldo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   9120
      TabIndex        =   19
      Top             =   7800
      Width           =   1575
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
      TabIndex        =   14
      Top             =   480
      Width           =   2175
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
      TabIndex        =   10
      Top             =   8160
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
      TabIndex        =   8
      Top             =   7800
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
      TabIndex        =   6
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dd/mm/aaaa"
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
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Menu ldso23 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tconteoo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim xproducto As String
Dim dbconteo As New ADODB.Recordset

Private Sub bodega_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Command1.Enabled = False Then
        DBGrid2.SetFocus
        Exit Sub

    End If

    If Len(bodega) = 0 Then
        bodega.SetFocus
        Exit Sub

    End If

    vendedor.SetFocus

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        Frame1.Visible = False
        Frame1.Enabled = False
        producto.SetFocus
        Exit Sub

    End If

    Command8_Click

End Sub

Private Sub Command1_Click()

    Dim found As Integer

    Dim sdx   As Double

    Dim vr

    Dim Tmp As String

    On Error GoTo cmd43_err

    If Len(fecha) = 0 Then
        fecha.SetFocus
        Exit Sub

    End If

    If Len(fecha) <> 10 Then
        fecha.SetFocus
        Exit Sub

    End If

    If Not IsDate(fecha) Then
        fecha.SetFocus
        Exit Sub

    End If

    If modelo = "ADICIONA" Then
        'MsgBox modelo
        'numero = busca_parame(0)
        'ax:
        '   found = valida_numero()
        '   If found = 1 Then
        '      sdx = Val(numero) + 1
        '      numero = "" & sdx
        '      GoTo ax
        '   End If
        found = grabar1()
        '   tmp = busca_parame(1)
        habilita 0
        habilita1 1

    End If
   
    If dbconteo.State = 1 Then dbconteo.Close
    dbconteo.Open "select * from conteofi where numero='" & Numero & "' order by hora ", cn, adOpenDynamic, adLockOptimistic
   
    Set DBGrid2.DataSource = dbconteo
    suma_sobrantes
               
    'dbGrid2.Col = 0
    'dbGrid2.Row = dbGrid2.VisibleRows - 1
    'dbGrid2.SetFocus
    Exit Sub
cmd43_err:
    MsgBox "Seleccione datos " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub Command11_Click()
    Frame3.Visible = True

End Sub

Private Sub Command13_Click()
    Frame6.Visible = True
    borra_forma
    producto.SetFocus

End Sub

Private Sub Command2_Click()
    suma_sobrantes

End Sub

Private Sub Command3_Click()

    If modelo = "ADICIONA" Then
        habilita 1
        habilita1 0
        Command1.SetFocus

    End If

End Sub

Private Sub Command4_Click()

    Dim buf As String

    buf = "select * from conteofi where NUMERO='" & Numero & "' and local='" & local1 & "' and fecha='" & fecha & "' and bodega='" & extra_loquesea(bodega) & "'"

    If Combo2 <> "%" Then
        buf = buf & " and " & Combo2 & " like '" & Text1 & "'"

    End If

    If Combo3 <> "%" Then
        If Combo3 = "PRODUCTO" Then
            buf = buf & " order by str(producto)"
        Else
            buf = buf & " order by " & Combo3

        End If

    End If

    If dbconteo.State = 1 Then dbconteo.Close
    dbconteo.Open buf, cn, adOpenDynamic, adLockOptimistic
    Set DBGrid2.DataSource = dbconteo
    suma_sobrantes

End Sub

Private Sub Command6_Click()

    Dim saldoini As Double

    Dim saldoant As Double

    Dim sobrante As Double

    Dim faltante As Double

    Dim found    As Integer

    On Error GoTo cmd5678_err

    suma_xx

    If Val(saldo) = Val("" & conteo) * Val("" & factor) + Val("" & cantidad1) Then
        MsgBox "Cantidades iguales no se actualizan", 48, "Aviso"
        Exit Sub

    End If

    If Not IsNumeric(conteo) Then
        Exit Sub
            
    End If

    'If Not IsNumeric(cantidad1) Then
    '   Exit Sub
    'End If
            
    If Not IsNumeric(factor) Then
        Exit Sub

    End If

    If Len(unidad) = 0 Then
        Exit Sub

    End If

    If Len(descripcio) = 0 Then
        Exit Sub

    End If

    If Len(producto) = 0 Then
        Exit Sub

    End If
            
    saldoini = Val("" & saldo)
    saldoant = Val("" & conteo) * Val("" & factor) + Val("" & cantidad1)
    sobrante = 0
    faltante = 0

    If saldoini = saldoant Then  'igual

    End If

    If saldoini < saldoant Then  'sobrante
        sobrante = Abs(saldoini - saldoant)

    End If

    If saldoini > saldoant Then  'faltante
        faltante = Abs(saldoini - saldoant)

    End If

    dbconteo.AddNew
    dbconteo.Fields("numero") = "" & Numero
    dbconteo.Fields("producto") = "" & producto
    dbconteo.Fields("descripcio") = Mid$("" & descripcio, 1, 60)
    dbconteo.Fields("unidad") = "" & unidad
    dbconteo.Fields("factor") = "" & factor
    dbconteo.Fields("linea") = "" & linea
    dbconteo.Fields("bodega") = extra_loquesea(bodega)
    dbconteo.Fields("vendedor") = extra_loquesea(vendedor)
    dbconteo.Fields("local") = local1
    dbconteo.Fields("FECHA") = fecha
    dbconteo.Fields("hora") = Format(Now, "hh:mm:ss")
    dbconteo.Fields("cantidad") = Val(saldo)
    dbconteo.Fields("saldoant") = Val(conteo)
    dbconteo.Fields("cantidad1") = Val(cantidad1)
    dbconteo.Fields("sobrante") = Val(sobrante)
    dbconteo.Fields("faltante") = Val(faltante)
    dbconteo.Update
            
    'aqui deberia actualizar en kardex
    'xant = Val("" & conteo)
    'If MsgBox("Actualizar Kardex", 1, "Aviso") = 1 Then
    found = grabarx(Val(saldo))
    'End If
    borra_forma
    producto.SetFocus
    Exit Sub
cmd5678_err:
    MsgBox "Aviso en grabar " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command7_Click()
    Frame6.Visible = False

End Sub

Private Sub Command8_Click()
    ejecuta 1

End Sub

Sub ejecuta(sw As Integer)

    Dim rconsulta As New ADODB.Recordset

    Dim cad       As String

    'If opcion1 = "1" Then  'bodega
    If Len(buffer) = 0 Then
        cad = "select Producto.Descripcio,Producto.producto,Producto.Marca,Producto.Familia,Producto.Subfamilia,Linea from producto   "

    End If

    If Len(buffer) > 0 Then
        cad = "select Producto.Descripcio,Producto.producto,Producto.Marca,Producto.Familia,Producto.Subfamilia,Linea from producto      where  " & Combo1 & " like '" & buffer & "%'"

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
    'End If

End Sub

Private Sub conteo_KeyPress(KeyAscii As Integer)

    If Len("" & linea) > 0 Then
        Exit Sub

    End If

    If KeyAscii <> 13 Then Exit Sub
    Command6_Click

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        producto = dbGrid1.columns(1)
        Frame1.Visible = False
        Frame1.Enabled = False
        producto.SetFocus
        producto_KeyPress 13

    End If

End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)

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

Private Sub dbgrid1_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim found As Integer

    'Exit Sub
End Sub

Private Sub fecha_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If Len(fecha) = 0 Then
        fecha = Format(Now, "dd/mm/yyyy")

    End If

    If KeyAscii = 27 Then
        Exit Sub

    End If

    If Len(fecha) <> 10 Then
        fecha.SetFocus
        Exit Sub

    End If

    If Not IsDate(fecha) Then
        fecha.SetFocus
        Exit Sub

    End If

    vendedor.SetFocus

End Sub

Private Sub fkli3e3_Click()

End Sub

Private Sub Form_Activate()

    Dim mytablex As New ADODB.Recordset

    If modelo = "MODIFICA" Then
        habilita 0
        habilita1 1
        Command1_Click
        'MsgBox "h"
        Exit Sub

    End If

    If modelo = "SOLO VER" Then
        habilita 1
        habilita1 1
        Command1_Click
        DBGrid2.AllowUpdate = False
        DBGrid2.Enabled = True
        Command4.Enabled = True
        Command2.Enabled = True
        Exit Sub

    End If

    If modelo = "ADICIONA" Then
        fecha = Format(Now, "dd/mm/yyyy")

    End If

    vendedor.Clear
    bodega.Clear
    'tipoi.Clear
    mytablex.Open "SELECT * FROM bodega", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        bodega.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("Nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    bodega.ListIndex = 0
    mytablex.Open "SELECT * FROM vendedor", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        vendedor.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("Nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    vendedor.ListIndex = 0

End Sub

Private Sub Form_Load()

    Frame1.Top = 10: Frame1.Left = 10
    '---------------------------

    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "DESCRIPCIO"
    Combo2.AddItem "PRODUCTO"
    Combo2.AddItem "FAMILIA"
    Combo2.AddItem "LINEA"
    Combo2.ListIndex = 1

    Combo3.Clear
    Combo3.AddItem "%"
    Combo3.AddItem "DESCRIPCIO"
    Combo3.AddItem "PRODUCTO"
    Combo3.AddItem "FAMILIA"
    Combo3.AddItem "LINEA"
    Combo3.ListIndex = 1
    opcion5 = 0

End Sub

Private Sub Label17_Click()
    Label20.Visible = False
    Label20.Caption = ""
    Frame4.Visible = False

End Sub

Private Sub Label20_Click()

    If Label20.Visible = True Then
        Label20.Visible = False

    End If

End Sub

Private Sub Label31_Click()
    borra_forma
    producto.SetFocus

End Sub

Private Sub ldso23_Click()

    If Label20.Visible = True Then
        Label20.Visible = False
        Exit Sub

    End If

    If Frame3.Visible = True Then
        Frame3.Visible = False
        Exit Sub

    End If

    If Frame1.Visible = True Then
        Frame1.Visible = False
        DBGrid2.Enabled = True
        DBGrid2.SetFocus
        Exit Sub

    End If

    If Command1.Enabled = False Then
   
        If modelo = "ADICIONA" Then
            habilita 1
            habilita1 0
            Command1.Enabled = True
            Command1.SetFocus
            Exit Sub

        End If

    End If

    tconteoo.Hide
    Unload tconteoo

End Sub

Function carga_fecha()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM bodega where codigo='" & extra_loquesea(bodega) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        fecha = "" & mytablex.Fields("fecha")
        carga_fecha = 1

    End If

    mytablex.Close

End Function

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
    Command13.Enabled = xsw
    Command4.Enabled = xsw
    DBGrid2.Enabled = xsw

    'dk78231.Enabled = xsw
End Sub

Sub habilita1(sw As Integer)

    Dim xsw

    If sw = 0 Then
        xsw = True

    End If

    If sw = 1 Then
        xsw = False

    End If

    Numero.Enabled = xsw
    Command1.Enabled = xsw
    local1.Enabled = xsw
    fecha.Enabled = xsw
    bodega.Enabled = xsw
    vendedor.Enabled = xsw

End Sub

Function busca_linea(buf As String)

    Dim mytablex As New ADODB.Recordset

    nt1 = ""
    nt2 = ""
    nt3 = ""
    nt4 = ""
    nt5 = ""
    nt6 = ""
    nt7 = ""
    nt8 = ""
    nt9 = ""
    nt10 = ""
    nt11 = ""
    nt12 = ""
    nt13 = ""
    nt14 = ""
    nt15 = ""
    nt16 = ""
    mytablex.Open "select * from linea where linea='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

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

End Sub

Private Sub sum823_Click()

End Sub

Private Sub producto_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    Dim abuf  As String

    If KeyAscii = 27 Then
        Command7_Click
        Exit Sub

    End If

    producto = Trim(producto)

    If KeyAscii <> 13 Then Exit Sub

    If Len(producto) = 0 Then
        producto.SetFocus
        Exit Sub

    End If

    found = busca_productof("" & producto)

    If found = 0 Then
        MsgBox "No existe producto ", 48, "Aviso"
        borra_forma
        producto.SetFocus
        Exit Sub

    End If

    conteo.SetFocus

End Sub

Private Sub producto_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        menu_productos

    End If

End Sub

Private Sub t1_Change()
    suma_xx

End Sub

Sub suma_xx()

    Dim sdx As Double

    If Len(linea) > 0 Then
        sdx = Val(t1) + Val(t2) + Val(t3) + Val(t4) + Val(t5) + Val(t6) + Val(t7) + Val(t8) + Val(t9) + Val(t10) + Val(t11) + Val(t12) + Val(t13) + Val(t14) + Val(t15) + Val(t16)
        sumax = Format(sdx, "0")
        sdx = Val(t1) + Val(t2) + Val(t3) + Val(t4) + Val(t5) + Val(t6) + Val(t7) + Val(t8) + Val(t9) + Val(t10) + Val(t11) + Val(t12) + Val(t13) + Val(t14) + Val(t15) + Val(t16)
        conteo = sdx

    End If

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
    t1 = "" & DBGrid2.columns(8)
    t2 = "" & DBGrid2.columns(9)
    t3 = "" & DBGrid2.columns(10)
    t4 = "" & DBGrid2.columns(11)
    t5 = "" & DBGrid2.columns(12)
    t6 = "" & DBGrid2.columns(13)
    t7 = "" & DBGrid2.columns(14)
    t8 = "" & DBGrid2.columns(15)
    t9 = "" & DBGrid2.columns(16)
    t10 = "" & DBGrid2.columns(17)
    t11 = "" & DBGrid2.columns(18)
    t12 = "" & DBGrid2.columns(19)
    t13 = "" & DBGrid2.columns(20)
    t14 = "" & DBGrid2.columns(21)
    t15 = "" & DBGrid2.columns(22)
    t16 = "" & DBGrid2.columns(23)

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

Function busca_equiva(buf As String) As Integer

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM productb where barras='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        buf = "" & mytablex.Fields("producto")
        busca_equiva = 1
        mytablex.Close
        Exit Function

    End If

    mytablex.Close
   
    mytablex.Open "SELECT * FROM producto where barras='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        buf = "" & mytablex.Fields("producto")
        busca_equiva = 1

    End If

    mytablex.Close

End Function

Function menu_productos()

    Dim cad       As String

    Dim rconsulta As New ADODB.Recordset

    cad = "SELECT * FROM producto  "

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True Or rconsulta.BOF = True Then
        rconsulta.Close
        Exit Function

    End If

    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.ListIndex = 0
    opcion1 = 1
    Frame1.Enabled = True
    Frame1.Visible = True
    buffer = ""
    buffer.SetFocus
    Command8_Click

End Function

Sub ir_primero1()

    On Error GoTo cmd771222_err

    dbconteo.MoveFirst
    'Data1.Refresh

    Exit Sub
cmd771222_err:
    Exit Sub

End Sub

Sub suma_sobrantes()

    On Error GoTo cmd45_err

    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    wsaldo = ""
    wconteo = ""
    wsobrante = ""
    wfaltante = ""
    ir_primero1

    Do

        If dbconteo.EOF Then Exit Do
        suma1 = suma1 + Val("" & dbconteo.Fields("cantidad"))
        suma2 = suma2 + Val("" & dbconteo.Fields("saldoant"))
        suma3 = suma3 + Val("" & dbconteo.Fields("sobrante"))
        suma4 = suma4 + Val("" & dbconteo.Fields("faltante"))
        dbconteo.MoveNext
    Loop
    wsaldo = "" & suma1
    wconteo = "" & suma2
    wsobrante = "" & suma3
    wfaltante = "" & suma4
    ir_ultimo
    Exit Sub
cmd45_err:
    MsgBox "Seleccione datos " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub ir_ultimo()

    On Error GoTo cmd4_err

    dbconteo.MoveLast
    Exit Sub
cmd4_err:
    Exit Sub

End Sub

Function grabar_almacen()

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from almacen where local='" & local1 & "' and producto='" & producto & "' and bodega='" & extra_loquesea(bodega) & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        mytablex.Fields("saldo") = Val("" & conteo) * Val("" & factor) + Val("" & cantidad1)
        mytablex.Update
    Else
        mytablex.AddNew
        mytablex.Fields("producto") = "" & producto
        mytablex.Fields("local") = local1
        mytablex.Fields("bodega") = extra_loquesea(bodega)
        mytablex.Fields("saldo") = Val("" & conteo) * Val("" & factor) + Val("" & cantidad1)
        mytablex.Update

    End If

    mytablex.Close

End Function

Function valida_numero()

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from cconteof where numero='" & Numero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        valida_numero = 1

    End If

    mytablex.Close

End Function

Function grabar1()

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from cconteof where numero='" & Numero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.AddNew
        'mytablex.Fields("numero") = numero
        mytablex.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
        mytablex.Fields("vendedor") = extra_loquesea(vendedor)
        mytablex.Fields("local") = local1
        mytablex.Fields("bodega") = extra_loquesea(bodega)
        mytablex.Update
    Else
        'mytablex.Edit
        'mytablex.Fields("numero") = numero
        mytablex.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
        mytablex.Fields("vendedor") = extra_loquesea(vendedor)
        mytablex.Fields("local") = local1
        mytablex.Fields("bodega") = extra_loquesea(bodega)
        mytablex.Update

    End If

    Numero = "" & mytablex.Fields("numero")
    mytablex.Close

End Function

Private Sub vendedor_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    Command1.SetFocus

End Sub

Function busca_parame(sw As Integer) As Double

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from parame where codigo='01'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If sw = 0 Then
            sdx = Val("" & mytablex.Fields("conteo")) + 1
            busca_parame = "" & sdx

        End If

        If sw = 1 Then
            'mytablex.Edit
            mytablex.Fields("conteo") = Numero
            mytablex.Update

        End If

    End If

    mytablex.Close

End Function

Function verifica_doble(xlocal As String, _
                        xproducto As String, _
                        xbodega As String, _
                        xfecha As String)

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from conteofi where numero='" & Numero & "' and local='" & xlocal & "' and producto='" & xproducto & "' and bodega='" & xbodega & "' and fecha='" & xfecha & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        verifica_doble = 1

    End If

    mytablex.Close

End Function

Function recalculo_saldos1(xbuf As String)

    Dim saldoini As Double

    Dim signo    As Double

    Dim buf      As String

    Dim sdx      As Double

    Dim found    As Integer

    Dim fechai   As String

    On Error GoTo cmd333_err

    fechai = Format(busca_paramee(extra_loquesea(bodega)), "dd/mm/yyyy")

    If Not IsDate(fechai) Then
        Exit Function

    End If

    found = kardexactualiza(local1, "" & xbuf, extra_loquesea(bodega), fechai, fecha)
    recalculo_saldos1 = 5
    Exit Function
cmd333_err:
    MsgBox "Aviso en Recalculo saldo 1" + error$, 48, "Aviso"
    Exit Function

End Function

Function busca_paramee(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM bodega where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_paramee = "" & mytablex.Fields("fecha")

    End If

    mytablex.Close

End Function

Function grabarx(saldoa As Double)

    Dim found    As Integer

    Dim acu      As String

    Dim sw       As Integer

    Dim xingreso As Double

    Dim xegreso  As Double

    On Error GoTo cmd400_err

    sw = 0
    xingreso = 0
    xegreso = 0
    acu = ""
    'MsgBox saldoa
    found = grabar_almacen()

    'If Val("" & mytabley.Fields("cantidad")) > Val("" & mytabley.Fields("saldoant")) Then
    If saldoa > Val("" & "" & conteo) * Val("" & "" & factor) + Val("" & cantidad1) Then
        xegreso = saldoa - Val("" & "" & conteo) * Val("" & "" & factor) - Val("" & cantidad1)
        acu = "T"  'salida
        found = graba_kardex(acu, xingreso, xegreso)

        'Exit Function
    End If

    If saldoa < Val("" & "" & conteo) * Val("" & "" & factor) Then
        xingreso = -saldoa + Val("" & "" & conteo) * Val("" & "" & factor) + Val("" & cantidad1)
        acu = "S"
        found = graba_kardex(acu, xingreso, xegreso)

        'Exit Function
    End If
   
    Exit Function
cmd400_err:
    MsgBox "grabarx " + error$, 48, "Aviso"
    Exit Function

End Function

Function graba_kardex(acu As String, xingreso As Double, xegreso As Double)

    On Error GoTo cmd781_err

    Dim mytablez As New ADODB.Recordset

    If mytablez.State = 1 Then mytablez.Close
    mytablez.Open "select * from detalle where tipo='90' ", cn, adOpenDynamic, adLockOptimistic

    mytablez.AddNew
    mytablez.Fields("estado") = "2"
    mytablez.Fields("acu") = acu

    If acu = "S" Then
        mytablez.Fields("tipo") = "E"
        mytablez.Fields("cantidad") = xingreso

    End If

    If acu = "T" Then
        mytablez.Fields("cantidad") = xegreso
        mytablez.Fields("tipo") = "S"

    End If

    mytablez.Fields("local") = local1
    mytablez.Fields("serie") = ""
    mytablez.Fields("numero") = "CF" & Numero
    mytablez.Fields("tipoclie") = "V"
    mytablez.Fields("codigo") = "OFICINA"
    mytablez.Fields("acu1") = ""
    'mytablez.Fields("fecha") = Format(Now, "dd/mm/yyyy")
    mytablez.Fields("moneda") = "S"
    mytablez.Fields("producto") = "" & producto
    mytablez.Fields("descripcio") = "" & descripcio
    mytablez.Fields("unidad") = "UND"
    mytablez.Fields("factor") = 1 'Val("" & mytabley.Fields("factor"))
    mytablez.Fields("precio") = 0
    mytablez.Fields("igv") = 19
    mytablez.Fields("neto") = 0
    mytablez.Fields("descuento") = 0
    mytablez.Fields("subtotal") = 0
    mytablez.Fields("impuesto") = 0
    mytablez.Fields("total") = 0
    mytablez.Fields("fecha") = Format("" & fecha, "dd/mm/yyyy")
    mytablez.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
    mytablez.Fields("hora") = Format(Now, "hh:mm:ss")
    mytablez.Fields("vendedor") = ""
    mytablez.Fields("bodega") = "" & extra_loquesea(bodega)
    mytablez.Fields("bodegaf") = ""
    mytablez.Fields("deslipo") = 0
    mytablez.Fields("flage") = ""
    mytablez.Fields("linea") = "" & linea
    mytablez.Fields("t1") = 0
    mytablez.Fields("t2") = 0
    mytablez.Fields("t3") = 0
    mytablez.Fields("t4") = 0
    mytablez.Fields("t5") = 0
    mytablez.Fields("t6") = 0
    mytablez.Fields("t7") = 0
    mytablez.Fields("t8") = 0
    mytablez.Fields("t9") = 0
    mytablez.Fields("t10") = 0
    mytablez.Fields("t11") = 0
    mytablez.Fields("t12") = 0
    mytablez.Fields("t13") = 0
    mytablez.Fields("t14") = 0
    mytablez.Fields("t15") = 0
    mytablez.Fields("t16") = 0
    mytablez.Fields("l1") = ""
    mytablez.Fields("l2") = ""
    mytablez.Fields("l3") = ""
    mytablez.Fields("l4") = ""
    'mytablez.Fields("local") = ""
    mytablez.Fields("proveedorp") = ""
    mytablez.Fields("observa1") = ""
    mytablez.Fields("observa2") = ""
    mytablez.Fields("observa3") = ""
    mytablez.Fields("observa4") = ""
    mytablez.Fields("zona") = ""
    mytablez.Fields("isc") = 0
    mytablez.Fields("tax") = 0
    mytablez.Fields("vtaneta") = 0
    mytablez.Fields("tcosto") = 0
    mytablez.Fields("ganancia") = 0
    mytablez.Fields("comision") = 0
    mytablez.Fields("usuario") = ""
    mytablez.Fields("caja") = ""
    mytablez.Fields("turno") = ""
    mytablez.Fields("servicio") = ""
    mytablez.Fields("comanda") = ""
    mytablez.Fields("mesa") = ""
    mytablez.Fields("salon") = ""
    mytablez.Fields("mesero") = ""
    'mytablez.Fields("local") = extra_loquesea(local1)
    'MsgBox "x"
    mytablez.Update
    mytablez.Close
    graba_kardex = 1
    Exit Function
cmd781_err:
    MsgBox "Error " + error$, 48, "Aviso"
    Exit Function

End Function

Function busca_productof(buf As String)

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim found    As Integer

    Dim buf1     As String

    Dim I        As Integer

    Dim ssw      As Integer

    Dim sw       As Integer

    I = 0

    found = 0

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM producto where producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        found = busca_equiva(buf) 'busca en la table codigo barras

        If found = 0 Then
            Exit Function

        End If

        mytablex.Open "SELECT * FROM producto where producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.Close
            Exit Function

        End If

    End If
      
    found = recalculo_saldos1("" & mytablex.Fields("producto"))

    If found = 0 Then
        Exit Function

    End If

    producto = "" & mytablex.Fields("producto")
    descripcio = Mid$("" & mytablex.Fields("descripcio"), 1, 60)
    unidad = "" & mytablex.Fields("unidad")
    factor = "" & mytablex.Fields("factor")
    linea = "" & mytablex.Fields("linea")
    mytablex.Close
   
    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM almacen where local='" & local1 & "' and producto='" & buf & "' and bodega='" & extra_loquesea(bodega) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        saldo = Val("" & mytablex.Fields("saldo")) '

    End If

    found = busca_linea(buf)
    mytablex.Close
    busca_productof = 1

End Function

Sub borra_forma()
    cantidad1 = ""
    descripcio = ""
    unidad = ""
    factor = ""
    linea = ""
    producto = ""
    saldo = ""
    conteo = ""
    nlinea = ""
   
    nt1 = ""
    nt2 = ""
    nt3 = ""
    nt4 = ""
    nt5 = ""
    nt6 = ""
    nt7 = ""
    nt8 = ""
    nt9 = ""
    nt10 = ""
    nt11 = ""
    nt12 = ""
    nt13 = ""
    nt14 = ""
    nt15 = ""
    nt16 = ""

End Sub

Function ve_descarga(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM tipo where  tipo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe

        Select Case "" & mytablex.Fields("tipodoc")

            Case "A", "B", "C", "D", "G", "E", "F", "J", "K", "L", "M", "P", "N", "O", "S", "T"
                ve_descarga = 1

        End Select

    End If

    mytablex.Close

End Function

