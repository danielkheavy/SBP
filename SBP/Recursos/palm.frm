VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form palm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Orion-PedidoPocket"
   ClientHeight    =   3390
   ClientLeft      =   150
   ClientTop       =   120
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
      Height          =   1575
      Left            =   0
      TabIndex        =   74
      Top             =   1800
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   117
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   116
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   115
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   114
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   113
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   112
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "7"
         Height          =   255
         Index           =   6
         Left            =   1200
         TabIndex        =   111
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         Height          =   255
         Index           =   7
         Left            =   1560
         TabIndex        =   110
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   109
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Index           =   9
         Left            =   480
         TabIndex        =   108
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "."
         Height          =   255
         Index           =   10
         Left            =   840
         TabIndex        =   107
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   106
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "B"
         Height          =   255
         Index           =   12
         Left            =   480
         TabIndex        =   105
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
         Height          =   255
         Index           =   13
         Left            =   840
         TabIndex        =   104
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "D"
         Height          =   255
         Index           =   14
         Left            =   1200
         TabIndex        =   103
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "E"
         Height          =   255
         Index           =   15
         Left            =   1560
         TabIndex        =   102
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F"
         Height          =   255
         Index           =   16
         Left            =   1920
         TabIndex        =   101
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "G"
         Height          =   255
         Index           =   17
         Left            =   2280
         TabIndex        =   100
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "H"
         Height          =   255
         Index           =   18
         Left            =   2640
         TabIndex        =   99
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "I"
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   98
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "J"
         Height          =   255
         Index           =   20
         Left            =   480
         TabIndex        =   97
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "K"
         Height          =   255
         Index           =   21
         Left            =   840
         TabIndex        =   96
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "L"
         Height          =   255
         Index           =   22
         Left            =   1200
         TabIndex        =   95
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Index           =   23
         Left            =   1560
         TabIndex        =   94
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N"
         Height          =   255
         Index           =   24
         Left            =   1920
         TabIndex        =   93
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "O"
         Height          =   255
         Index           =   25
         Left            =   2280
         TabIndex        =   92
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P"
         Height          =   255
         Index           =   26
         Left            =   2640
         TabIndex        =   91
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Q"
         Height          =   255
         Index           =   27
         Left            =   120
         TabIndex        =   90
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         Height          =   255
         Index           =   28
         Left            =   480
         TabIndex        =   89
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         Height          =   255
         Index           =   29
         Left            =   840
         TabIndex        =   88
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "T"
         Height          =   255
         Index           =   30
         Left            =   1200
         TabIndex        =   87
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "U"
         Height          =   255
         Index           =   31
         Left            =   1560
         TabIndex        =   86
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "V"
         Height          =   255
         Index           =   32
         Left            =   1920
         TabIndex        =   85
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "W"
         Height          =   255
         Index           =   33
         Left            =   2280
         TabIndex        =   84
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X"
         Height          =   255
         Index           =   34
         Left            =   2640
         TabIndex        =   83
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Y"
         Height          =   255
         Index           =   35
         Left            =   120
         TabIndex        =   82
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Z"
         Height          =   255
         Index           =   36
         Left            =   480
         TabIndex        =   81
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DEL"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   37
         Left            =   2280
         TabIndex        =   80
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ESC"
         Height          =   255
         Index           =   38
         Left            =   1560
         TabIndex        =   79
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ENTER"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   39
         Left            =   2280
         TabIndex        =   78
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CLOSE"
         Height          =   255
         Index           =   40
         Left            =   2280
         TabIndex        =   77
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label xteclado 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   76
         Top             =   1680
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "*"
         Height          =   255
         Index           =   41
         Left            =   1920
         TabIndex        =   75
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFF00&
      Caption         =   "modifica"
      Height          =   2535
      Left            =   2640
      TabIndex        =   57
      Top             =   720
      Visible         =   0   'False
      Width           =   3255
      Begin VB.TextBox mcant 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   840
         MaxLength       =   5
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label mtotal 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   72
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label mpventa 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   70
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label mproducto 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   68
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label53 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   67
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Grabar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   66
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label51 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label50 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label48 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Factor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label47 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label munidad 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   61
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label mdescripcio 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   60
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label mfactor 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   59
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Ingrese su Nro Terminal"
      Height          =   3255
      Left            =   2520
      TabIndex        =   52
      Top             =   360
      Width           =   3255
      Begin VB.TextBox clave 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label terminal 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   56
         Top             =   2640
         Width           =   105
      End
      Begin VB.Label Label44 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vendedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   55
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Entrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1920
         TabIndex        =   54
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sistema Orion V5.0 Pocket Pc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   2880
         Width           =   3015
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFF00&
      Caption         =   "Pedidos"
      Height          =   2895
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   3255
      Begin VB.TextBox pedido 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   840
         MaxLength       =   11
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cerrar Programa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   50
         Top             =   2520
         Width           =   3255
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Enter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   49
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label41 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pedido"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFF00&
      Caption         =   "Productos"
      Height          =   2535
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
      Begin VB.TextBox xcodigo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   11
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Teclado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   40
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label30 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   735
      End
      Begin VB.Label xstock 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   38
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label xpventa 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   37
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label xfactor 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   36
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label xnombre 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   35
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label xunidad 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   34
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   33
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label28 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Stock"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pventa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Factor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   27
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   26
         Top             =   1200
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Busqueda de Productos"
      Height          =   3255
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
      Begin VB.TextBox buffer 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   11
         TabIndex        =   45
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox condicion 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   240
         Width           =   1815
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "palm.frx":0000
         Height          =   2415
         Left            =   120
         OleObjectBlob   =   "palm.frx":0014
         TabIndex        =   20
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   43
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Teclado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   41
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   22
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   21
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Finalizar el Pedido"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
      Begin VB.TextBox observa 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   30
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox nombre 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   60
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox Codigo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   11
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   17
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8520
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8520
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "palm.frx":1087
      Height          =   2415
      Left            =   0
      OleObjectBlob   =   "palm.frx":109B
      TabIndex        =   2
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label xpedido 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   51
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label40 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   0
      TabIndex        =   46
      Top             =   2880
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Modifica"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   42
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   4080
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ClearAll"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BrrLinea"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BuscaProd."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cod.Exa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   6000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Total 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "palm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xlocal As String
Dim xserie As String
Dim xtipo As String
Option Explicit
Sub abrir_gaveta()
Dim found As Integer
found = abre_puerto("LPT1", 1)
If found = 0 Then Exit Sub
MsgBox "Solo Demo basico-Proceso Realizado ", 48, "Aviso"
End Sub
Function abre_puerto(apuerto As String, sw As Integer)  'solo gaveta dinero
Dim buf As String
Dim i
On Error GoTo cmd1_err
Select Case apuerto
       Case "LPT1", "LPT2", "LPT3", "LPT4", "LPT5"
            If sw = 0 Then   'star
               'buf = Chr$(27) + "d" + Chr$(0) 'star sp342
               i = FreeFile
               Open apuerto For Output As i
               buf = Chr$(28) + Chr$(29)  'star sp200
               buf = Chr$(7)   'star
               Print #i, buf;
               Close i
            End If
            If sw = 1 Then   'epson
               i = FreeFile
               Open apuerto For Output As i
               buf = Chr$(27) + "i"  'epson
               buf = Chr$(27) + "p" + Chr$(0) + Chr$(25) + Chr$(250) 'EPSON
               Print #i, buf;
               Close i
            End If
       Case "COM1", "COM2", "COM3", "COM4", "COM5"
            If sw = 0 Then   'star
               Open apuerto For Input As #9
               'buf = Chr$(27) + "d" + Chr$(0) 'star sp342
               buf = Chr$(28) + Chr$(29)  'star sp200
               Print #9, buf;
            End If
            If sw = 1 Then   'epson
               Open apuerto For Input As #9
               buf = Chr$(27) + "i"  'epson
               Print #9, buf;
               Close #9
            End If
End Select
abre_puerto = 1
Exit Function
cmd1_err:
MsgBox "Error en abrir Gaveta", 48, "Aviso"
Exit Function
End Function

Private Sub buffer_KeyPress(KeyAscii As Integer)
'
End Sub

Private Sub buffer_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> 13 And KeyCode <> 27 Then
   consulta_productos
End If

End Sub


Function busca_terminal()
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("parameca")
mytablex.Index = "caja"
mytablex.Seek "=", "" & clave
If Not mytablex.NoMatch Then
   busca_terminal = 1
End If
'------------------------------------- ------------
mytablex.Close

End Function

Private Sub clave_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(clave) = 0 Then
   clave.SetFocus
   Exit Sub
End If
clave = UCase(clave)
'If Mid$(clave, 1, 1) <> "T" Then
'   clave = ""
'   MsgBox "No es Terminal ", 48, "Aviso"
'   clave.SetFocus
'   Exit Sub
'End If
terminal = ""
found = busca_vendedor()
If found = 0 And Len(terminal) = 0 Then
   clave = ""
   terminal = ""
   MsgBox "Vendedor no existe", 48, "Aviso"
   clave.SetFocus
   Exit Sub
End If
gusuario = "_w" & terminal
found = copiar_temporalp1()
If found = 0 Then
   MsgBox "Ya existe El uso ", 48, "Aviso"
End If

'validamos si esta prohibido entrara a estas partes

Set mytable11 = mydbxglo.OpenTable("parameca")
mytable11.Index = "caja"
mytable11.Seek "=", "" & terminal
If Not mytable11.NoMatch Then
End If
If mytable11.NoMatch Then
   clave = ""
   MsgBox "No existe Terminal", 48, "Aviso"
   mytable11.Close
   clave.SetFocus
   Exit Sub
End If
carga_inicial
Frame3.Visible = False
SQL_pedido
found = suma_detalle()
Frame5.Visible = False
Label41_Click
pedido.SetFocus
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(Codigo) > 0 Then
found = busca_codigo()
If found = 0 Then
End If
End If
nombre.SetFocus
End Sub

Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   Frame2.Visible = False
   Exit Sub
End If
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim found As Integer
If KeyCode = 13 Then
   found = adiciona_registro()
   If found = 0 Then
      MsgBox "No se Pudo adicionar ", 48, "Aviso"
      DBGrid1.SetFocus
      Exit Sub
   End If
   Frame5.Visible = False
   found = suma_detalle()
   Frame1.Visible = False
   Data2.Refresh
   found = ir_ultimo()
   If found = 0 Then
      Data2.Refresh
   End If
   DBGrid2.Col = 1
   DBGrid2.SetFocus
   
End If

End Sub
Function adiciona_registro()
On Error GoTo cmd56_err
Data2.Recordset.AddNew
Data2.Recordset.Fields("local") = xlocal
Data2.Recordset.Fields("serie") = xserie
Data2.Recordset.Fields("tipo") = xtipo
Data2.Recordset.Fields("fecha") = Format(Now, "dd/mm/yyyy")
Data2.Recordset.Fields("hora") = Format(Now, "hh:mm:ss")
Data2.Recordset.Fields("numero") = pedido
Data2.Recordset.Fields("producto") = "" & Data1.Recordset.Fields("producto")
Data2.Recordset.Fields("descripcio") = "" & Data1.Recordset.Fields("descripcio")
Data2.Recordset.Fields("unidad") = "" & Data1.Recordset.Fields("unidad1")
Data2.Recordset.Fields("factor") = Val("" & Data1.Recordset.Fields("factor1"))
Data2.Recordset.Fields("precio") = Val("" & Data1.Recordset.Fields("pventa1"))
Data2.Recordset.Fields("igv") = Val("" & Data1.Recordset.Fields("igv"))
Data2.Recordset.Fields("total") = Val("" & Data1.Recordset.Fields("pventa1"))
Data2.Recordset.Fields("cantidad") = 1
Data2.Recordset.Update
adiciona_registro = 1
Exit Function
cmd56_err:
Exit Function
End Function
Sub adiciona_registrox()
Data2.Recordset.AddNew
Data2.Recordset.Fields("producto") = "" & xcodigo
Data2.Recordset.Fields("descripcio") = "" & xnombre
Data2.Recordset.Fields("unidad") = "" & xunidad
Data2.Recordset.Fields("factor") = Val("" & xfactor)
Data2.Recordset.Fields("precio") = Val("" & xpventa)
Data2.Recordset.Fields("igv") = 19 'Val("" & xigv)
Data2.Recordset.Fields("total") = Val("" & xpventa)
Data2.Recordset.Fields("cantidad") = 1
Data2.Recordset.Update

End Sub
Function ir_ultimo()
On Error GoTo cmd90_err
Data2.Recordset.MoveLast
Exit Function
cmd90_err:
Exit Function
End Function

Private Sub DBGrid2_AfterColUpdate(ByVal ColIndex As Integer)
Dim found As Integer
Select Case ColIndex
       Case 1, 3
            found = suma_detalle()
            found = ir_ultimo()
            'DBGrid2.Col = 0
            'DBGrid2.Row = DBGrid2.VisibleRows - 1

            'Data2.Refresh
            
End Select
End Sub

Private Sub DBGrid2_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
Select Case ColIndex
       Case 0
            Cancel = True
            Exit Sub
       Case 2
            Cancel = True
            Exit Sub
End Select
End Sub

Private Sub DBGrid2_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Select Case ColIndex
       Case 0
            Cancel = True
            Exit Sub
       Case 1
            If Not IsNumeric(DBGrid2.Columns(1)) Then
               Cancel = True
               Exit Sub
            End If
            DBGrid2.Columns(2) = Val("" & DBGrid2.Columns(1)) * Val("" & DBGrid2.Columns(3))
       Case 3
            If Not IsNumeric(DBGrid2.Columns(3)) Then
               Cancel = True
               Exit Sub
            End If
            DBGrid2.Columns(2) = Val("" & DBGrid2.Columns(1)) * Val("" & DBGrid2.Columns(3))
End Select
End Sub

Private Sub DBGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H2E Then  'borrar linea
   borra_linea
End If
End Sub
Sub borra_linea()
        If DBGrid2.Row = -1 Then
        MsgBox "No hay ningún registro para eliminar", vbInformation
        Exit Sub
        End If
        If MsgBox("Desea Borrar ", 1, "Aviso") <> 1 Then
           DBGrid2.SetFocus
           Exit Sub
        End If
        Data2.Recordset.Delete
        Data2.Refresh

End Sub

Private Sub Form_Activate()
xtipo = "P"
xlocal = "01"
xserie = "P"
End Sub

Private Sub Form_Load()
Dim found As Integer
globaldir = "C:\orion.v5\001d\06"
'gusuario = "dproform"
'found = copiar_temporal()
'If found = 0 Then
'   MsgBox "Ya existe El uso ", 48, "Aviso"
'End If
Set mydbxglo = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
condicion.Clear
condicion.AddItem "Producto.Producto"
condicion.AddItem "Producto.Descripcio"
condicion.ListIndex = 0
'carga_inicial
xtipo = ""
xlocal = "01"
xserie = ""
Label44_Click

End Sub
Function suma_detalle()
Dim found As Integer
Dim xtotal As Double
Dim rs
'found = ir_inicio()
xtotal = 0
Set rs = Data2.Recordset.Clone
Do
If rs.EOF Then Exit Do
xtotal = xtotal + Val("" & rs.Fields("total"))
rs.MoveNext
Loop
Total = Format(xtotal, "0.00")
End Function
Function copiar_temporalp1()
On Error GoTo cmd23_err
borrar_kill gusuario & ".dbf"
borrar_kill gusuario & ".cdx"
FileCopy globaldir & "\tdetalle.dbf", globaldir & "\" & gusuario & ".dbf"
FileCopy globaldir & "\tdetalle.cdx", globaldir & "\" & gusuario & ".cdx"
copiar_temporalp1 = 1
Exit Function
cmd23_err:
MsgBox error$
Exit Function
End Function
Sub borrar_kill(buf)
On Error GoTo cmd67_err
Kill buf
Exit Sub
cmd67_err:
Exit Sub
End Sub

Sub SQL_pedido()
               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldir
               Data2.RecordSource = "select * from " & gusuario & " where local='" & xlocal & "' and serie='" & xserie & "' and tipo='" & xtipo & "' and numero='" & pedido & "'"
               Data2.Refresh

End Sub

Private Sub Label1_Click()
xteclado = "Codigo"
Frame5.Visible = True

End Sub


Private Sub Label10_Click()
'Dim sdx As Double
'Dim found As Integer
'Label40.Visible = True
'found = suma_detalle()
'Frame2.Visible = True
'Label1_Click
'Codigo.SetFocus
'esto se ha adicionado despues
Dim found As Integer
found = valida()
If found = 0 Then
   Exit Sub
End If
found = grabar()
If found = 0 Then Exit Sub
borrar_todo
inicializa_todo
Label40.Visible = False
Frame5.Visible = False
Label17_Click
Frame7.Visible = True
pedido = ""
Label41_Click
pedido.SetFocus
End Sub

Private Sub Label11_Click()
Dim found As Integer
Label40.Visible = False
found = suma_detalle()
Frame1.Visible = False
Data2.Refresh
Frame5.Visible = False

DBGrid2.SetFocus
End Sub

Private Sub Label12_Click()
consulta_productos
End Sub

Private Sub Label13_Click()
xteclado = "Nombre"
Frame5.Visible = True

End Sub

Private Sub Label14_Click()
xteclado = "Observa"
Frame5.Visible = True

End Sub

Private Sub Label16_Click()
observa_KeyPress 13
End Sub

Private Sub Label17_Click()
Dim found  As Integer
Label40.Visible = False
found = suma_detalle()
Data2.Refresh
Frame2.Visible = False
Label40.Visible = False
Frame5.Visible = False
DBGrid2.SetFocus
End Sub

Private Sub Label22_Click()
Dim found As Integer
Label40.Visible = False
found = suma_detalle()
Frame4.Visible = False
Frame5.Visible = False
Data2.Refresh
DBGrid2.SetFocus

End Sub

Private Sub Label23_Click()
xcodigo_KeyPress 13
End Sub

Private Sub Label29_Click()
Dim found As Integer
   adiciona_registrox
   found = suma_detalle()
   Frame4.Visible = False
   Data2.Refresh
   ir_ultimo
   Label40.Visible = False
   DBGrid2.SetFocus

End Sub

Private Sub Label3_Click()
Label40.Visible = True
Frame4.Visible = True
inicializa_xproducto
xcodigo.SetFocus
End Sub

Private Sub Label31_Click(Index As Integer)
If xteclado = "Observa" Then
   Select Case Index
          Case 39 'emter
               Frame5.Visible = False
               If Frame4.Visible = False Then Exit Sub
               observa_KeyPress 13
               Exit Sub
          Case 37 'delete
               observa = ""
               Exit Sub
          Case 40 'close
               Frame5.Visible = False
               If Frame4.Visible = False Then Exit Sub
               observa.SetFocus
               Exit Sub
          
               
   End Select
   observa = observa & Label31(Index)
   Exit Sub
End If

If xteclado = "Codigo" Then
   Select Case Index
          Case 39 'emter
               Frame5.Visible = False
               If Frame4.Visible = False Then Exit Sub
               Codigo_KeyPress 13
               Exit Sub
          Case 37 'delete
               Codigo = ""
               Exit Sub
          Case 40 'close
               Frame5.Visible = False
               If Frame4.Visible = False Then Exit Sub
               Codigo.SetFocus
               Exit Sub
          
               
   End Select
   Codigo = Codigo & Label31(Index)
   Exit Sub
End If
If xteclado = "Nombre" Then
   Select Case Index
          Case 39 'emter
               Frame5.Visible = False
               If Frame4.Visible = False Then Exit Sub
               nombre_KeyPress 13
               Exit Sub
          Case 37 'delete
               nombre = ""
               Exit Sub
          Case 40 'close
               Frame5.Visible = False
               If Frame4.Visible = False Then Exit Sub
               nombre.SetFocus
               Exit Sub
          
               
   End Select
   nombre = nombre & Label31(Index)
   Exit Sub
End If

If xteclado = "Producto" Then
   Select Case Index
          Case 39 'emter
               Frame5.Visible = False
               If Frame4.Visible = False Then Exit Sub
               xcodigo_KeyPress 13
               Exit Sub
          Case 37 'delete
               xcodigo = ""
               Exit Sub
          Case 40 'close
               Frame5.Visible = False
               If Frame4.Visible = False Then Exit Sub
               xcodigo.SetFocus
               Exit Sub
          
               
   End Select
   xcodigo = xcodigo & Label31(Index)
   Exit Sub
End If
If xteclado = "Pedido" Then
   Select Case Index
          Case 39 'emter
               Frame5.Visible = False
               If Frame7.Visible = False Then Exit Sub
               pedido_KeyPress 13
               Exit Sub
          Case 37 'delete
               pedido = ""
               Exit Sub
          Case 40 'close
               Frame5.Visible = False
               If Frame7.Visible = False Then Exit Sub
               pedido.SetFocus
               Exit Sub
   End Select
   pedido = pedido & Label31(Index)
End If
If xteclado = "Clave" Then
   Select Case Index
          Case 39 'emter
               Frame5.Visible = False
               If Frame3.Visible = False Then Exit Sub
               clave_KeyPress 13
               Exit Sub
          Case 37 'delete
               pedido = ""
               Exit Sub
          Case 40 'close
               Frame5.Visible = False
               If Frame3.Visible = False Then Exit Sub
               clave.SetFocus
               Exit Sub
   End Select
   clave = clave & Label31(Index)
End If
If xteclado = "Busca" Then
   Select Case Index
          Case 39 'emter
               Frame5.Visible = False
               If Frame1.Visible = False Then Exit Sub
               buffer_KeyPress 13
               
               Exit Sub
          Case 37 'delete
               buffer = ""
               Exit Sub
          Case 40 'close
               Frame5.Visible = False
               If Frame1.Visible = False Then Exit Sub
               buffer.SetFocus
               Exit Sub
   End Select
   buffer = buffer & Label31(Index)
   consulta_productos
End If
If xteclado = "Cant" Then
   Select Case Index
          Case 39 'emter
               Frame5.Visible = False
               If Frame8.Visible = False Then Exit Sub
               mcant_KeyPress 13
               Exit Sub
          Case 37 'delete
               mcant = ""
               Exit Sub
          Case 40 'close
               Frame5.Visible = False
               If Frame8.Visible = False Then Exit Sub
               mcant.SetFocus
               Exit Sub
   End Select
   mcant = mcant & Label31(Index)
End If

End Sub

Private Sub Label32_Click()
xteclado = "Producto"
Frame5.Visible = True

End Sub

Private Sub Label33_Click()
xteclado = "Busca"
Frame5.Visible = True

End Sub

Private Sub Label34_Click()
On Error GoTo cmd123_err
mproducto = "" & Data2.Recordset.Fields("producto")
mdescripcio = "" & Data2.Recordset.Fields("descripcio")
munidad = "" & Data2.Recordset.Fields("unidad")
mfactor = "" & Data2.Recordset.Fields("factor")
mpventa = "" & Data2.Recordset.Fields("precio")
mcant = "" '& Data2.Recordset.Fields("cantidad")
mtotal = "" & Data2.Recordset.Fields("total")
Frame8.Visible = True
Label47_Click
mcant.SetFocus
Exit Sub
cmd123_err:
Data2.Refresh
Exit Sub

End Sub

Private Sub Label35_Click()

End Sub

Private Sub Label36_Click()
xteclado = "Codigo"
Frame5.Visible = True

End Sub

Private Sub Label37_Click()

End Sub

Private Sub Label38_Click()

End Sub

Private Sub Label39_Click()
Dim found As Integer
   found = adiciona_registro()
   found = suma_detalle()
   Frame1.Visible = False
   'Data2.Refresh
   'found = ir_ultimo()
   DBGrid2.Col = 1
   DBGrid2.Row = DBGrid2.VisibleRows - 1
   Label40.Visible = False
   Frame5.Visible = False
   DBGrid2.SetFocus

End Sub

Private Sub Label4_Click()
Label40.Visible = True
buffer = ""
Frame1.Visible = True
consulta_productos
DBGrid1.SetFocus
Label33_Click

End Sub
Sub consulta_productos()



               Data1.Connect = "foxpro 2.5;"
               Data1.DatabaseName = globaldir
               Data1.RecordSource = "select Producto.descripcio,precios.pventa1,producto.producto,Precios.unidad1,Precios.factor1,producto.igv  from producto  left join precios on producto.producto=precios.producto  where precios.local='01' and  " & condicion & " like '" & buffer & "*'"
               Data1.Refresh
               DBGrid1.Columns(0).Width = 1800
               DBGrid1.Columns(1).Width = 500


End Sub
Sub borrar_todo()
Dim found As Integer
On Error GoTo cmd90_err
found = ir_inicio()
If found = 0 Then
   Data2.Refresh
   Exit Sub
End If
Do
If Data2.Recordset.EOF Then Exit Do
Data2.Recordset.Delete
Data2.Refresh
Loop
DBGrid2.SetFocus
Exit Sub
cmd90_err:
Exit Sub
End Sub
Function ir_inicio()
On Error GoTo cmd891_err
Data2.Recordset.MoveFirst
ir_inicio = 1
Exit Function
cmd891_err:
Exit Function
End Function

Private Sub Label41_Click()
xteclado = "Pedido"
Frame5.Visible = True

End Sub

Private Sub Label42_Click()
End Sub

Private Sub Label43_Click()
pedido_KeyPress 13
End Sub

Private Sub Label44_Click()
xteclado = "Clave"
Frame5.Visible = True

End Sub

Private Sub Label45_Click()

End Sub

Private Sub Label47_Click()
xteclado = "Cant"
Frame5.Visible = True

End Sub

Private Sub Label49_Click()
End Sub

Private Sub Label5_Click()
Dim found As Integer
borra_linea
found = suma_detalle()
Data2.Refresh
End Sub

Private Sub Label52_Click()
Dim found As Integer
On Error GoTo cmd8911_err
If Not IsNumeric(mcant) Then
   Exit Sub
End If
Data2.Recordset.Edit
Data2.Recordset.Fields("cantidad") = Val(mcant)
Data2.Recordset.Fields("TOTAL") = Val("" & Data2.Recordset.Fields("cantidad")) * Val("" & Data2.Recordset.Fields("PRECIO"))
Data2.Recordset.Update
Label40.Visible = False
Frame8.Visible = False
found = suma_detalle()
Frame5.Visible = False
Label49_Click
Exit Sub
cmd8911_err:
MsgBox "No se puede Grabar", 48, "Aviso"
Exit Sub

End Sub

Private Sub Label53_Click()
Frame5.Visible = False
Frame8.Visible = False
End Sub

Private Sub Label6_Click()
Dim found As Integer
borrar_todo
found = suma_detalle()
Frame7.Visible = True
pedido = ""
pedido.SetFocus


End Sub

Private Sub Label7_Click()
End
End Sub

Private Sub Label8_Click()
clave_KeyPress 13
End Sub

Private Sub mcant_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
End Sub

Private Sub nombre_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
observa.SetFocus

End Sub

Private Sub nombre_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   Codigo.SetFocus
   Exit Sub
End If
End Sub

Function busca_vendedor()
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("vendedor")
mytablex.Index = "codigo"
mytablex.Seek "=", "" & clave
If Not mytablex.NoMatch Then
   busca_vendedor = 1
   'MsgBox mytablex.Fields("pocket")
   terminal = "" & mytablex.Fields("pocket")
End If
'------------------------------------- ------------
mytablex.Close

End Function
Function grabar()
Dim found As Integer
'found = numero_libre()
If Len(Codigo) > 0 And Len(nombre) > 0 Then
found = adiciona_cliente()
End If
found = adiciona_cabeza()
found = adiciona_detalle()
grabar = 1
End Function
Function adiciona_cliente()
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("clientes")
mytablex.Index = "codigo"
mytablex.Seek "=", "" & Codigo
If Not mytablex.NoMatch Then
   mytablex.Edit
   mytablex.Fields("codigo") = "" & Codigo
   mytablex.Fields("nombre") = "" & nombre
   mytablex.Update
End If
If mytablex.NoMatch Then
   mytablex.AddNew
   mytablex.Fields("codigo") = "" & Codigo
   mytablex.Fields("nombre") = "" & nombre
   mytablex.Update
End If
mytablex.Close
End Function
Function valida()
Dim found As Integer
'If Len(Codigo) = 0 Then
'   Codigo.SetFocus
'   Exit Function
'End If
'If Len(nombre) = 0 Then
'   Label13_Click
'   nombre.SetFocus
'   Exit Function
'End If
'found = busca_codigo()
'If found = 0 Then
'   Codigo.SetFocus
'   Exit Function
'End If
'If Len(vendedor) = 0 Then
'   vendedor.SetFocus
'   Exit Function
'End If
'
'found = busca_vendedor()
'If found = 0 Then
'   vendedor.SetFocus
'   Exit Function
'End If
valida = 1

End Function
Function numero_libre()
'Dim found As Integer
'Dim sdx As Double
'sdx = Val("" & numero)
'Set mytablex = mydbxglo.OpenTable("cproform")
'mytablex.Index = "tfactura"
'recibe:
'mytablex.Seek "=", xlocal, xtipo, xserie, pedido
'If Not mytablex.NoMatch Then
'   sdx = sdx + 1
'   numero = "" & sdx
'   GoTo recibe
'End If

End Function
Function adiciona_cabeza()
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("cproform")
mytablex.Index = "tfactura"
mytablex.Seek "=", xlocal, xtipo, xserie, pedido
If mytablex.NoMatch Then
   mytablex.AddNew
   mytablex.Fields("local") = "" & xlocal
   mytablex.Fields("tipo") = "" & xtipo
   mytablex.Fields("serie") = "" & xserie
   mytablex.Fields("numero") = "" & pedido
   mytablex.Fields("codigo") = "" & Codigo
   mytablex.Fields("nombre") = "" & nombre
   mytablex.Fields("moneda") = "" & mytable11.Fields("moneda")
   mytablex.Fields("bodega") = "" & mytable11.Fields("bodega")
   mytablex.Fields("estado") = "2"
   mytablex.Fields("usuario") = clave 'extra_loquesea("" & vendedor)
   mytablex.Fields("total") = Val("" & Total)
   mytablex.Fields("CAJA") = "" & terminal
   mytablex.Fields("tipoclie") = "C"
   mytablex.Fields("ACU") = "G"
   mytablex.Fields("SERVICIO") = "*"
   mytablex.Fields("vendedor") = clave 'extra_loquesea("" & vendedor)
   mytablex.Fields("observa") = "" & observa
   mytablex.Fields("fechaCREA") = Format(Now, "dd/mm/yyyy")
   mytablex.Fields("fecha") = Format(Now, "dd/mm/yyyy")
   mytablex.Fields("fechaE") = Format(Now, "dd/mm/yyyy")
   mytablex.Fields("hora") = Format(Now, "hh:mm:ss")
   mytablex.Update
End If
If Not mytablex.NoMatch Then
   mytablex.Edit
   mytablex.Fields("local") = "" & xlocal
   mytablex.Fields("tipo") = "" & xtipo
   mytablex.Fields("serie") = "" & xserie
   mytablex.Fields("numero") = "" & pedido
   mytablex.Fields("codigo") = "" & Codigo
   mytablex.Fields("nombre") = "" & nombre
   mytablex.Fields("moneda") = "" & mytable11.Fields("moneda")
   mytablex.Fields("bodega") = "" & mytable11.Fields("bodega")
   mytablex.Fields("estado") = "2"
   mytablex.Fields("usuario") = clave 'extra_loquesea("" & vendedor)
   mytablex.Fields("total") = Val("" & Total)
   mytablex.Fields("CAJA") = "" & terminal
   mytablex.Fields("tipoclie") = "C"
   mytablex.Fields("ACU") = "G"
   mytablex.Fields("SERVICIO") = "*"
   mytablex.Fields("vendedor") = clave 'extra_loquesea("" & vendedor)
   mytablex.Fields("observa") = "" & observa
   mytablex.Fields("fechaCREA") = Format(Now, "dd/mm/yyyy")
   mytablex.Fields("fecha") = Format(Now, "dd/mm/yyyy")
   mytablex.Fields("fechaE") = Format(Now, "dd/mm/yyyy")
   mytablex.Fields("hora") = Format(Now, "hh:mm:ss")
   mytablex.Update
End If
'------------------------------------- ------------
mytablex.Close
End Function
Function adiciona_detalle()
Dim i As Integer
Dim found As Integer
Dim mytablex As Table
Dim rs
Set mytablex = mydbxglo.OpenTable("dproform")
mytablex.Index = "tdetalle"
apx1:
mytablex.Seek "=", xlocal, xtipo, xserie, pedido
If Not mytablex.NoMatch Then
   mytablex.Delete
   GoTo apx1
End If
Set rs = Data2.Recordset.Clone
Do
If rs.EOF Then Exit Do
   mytablex.AddNew
   For i = 0 To rs.Fields.Count - 1
   mytablex.Fields(i) = rs.Fields(i)
   Next i
   graba_detalle mytablex
   mytablex.Update
   rs.MoveNext
Loop
mytablex.Close
End Function
Sub graba_detalle(mytablex As Table)
   mytablex.Fields("local") = "" & xlocal
   mytablex.Fields("tipo") = "" & xtipo
   mytablex.Fields("serie") = "" & xserie
   mytablex.Fields("numero") = "" & pedido
   mytablex.Fields("codigo") = "" & Codigo
   'mytablex.Fields("nombre") = "" & nombre
   mytablex.Fields("vendedor") = clave 'extra_loquesea("" & vendedor)
   mytablex.Fields("fecha") = Format(Now, "dd/mm/yyyy")
   mytablex.Fields("hora") = Format(Now, "hh:mm:ss")
   mytablex.Fields("moneda") = "" & mytable11.Fields("moneda")
   mytablex.Fields("bodega") = "" & mytable11.Fields("bodega")
   mytablex.Fields("estado") = "2"
   mytablex.Fields("usuario") = clave ' extra_loquesea("" & vendedor)
End Sub
Sub inicializa_todo()
Dim found As Integer
Codigo = ""
nombre = ""
'vendedor.ListIndex = 0
terminal = ""
observa = ""
'cerrar_data
'found = copiar_temporalp1()
'If found = 0 Then
'   MsgBox "Ya existe El uso ", 48, "Aviso"
'End If
'SQL_pedido
found = suma_detalle()
DBGrid2.SetFocus
End Sub
Sub cerrar_data()
On Error GoTo cmd8911_err
Data2.Refresh
Data2.Recordset.Close
Exit Sub
cmd8911_err:
'MsgBox "x"
Exit Sub
End Sub


Private Sub observa_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
found = valida()
If found = 0 Then
   Exit Sub
End If
found = grabar()
If found = 0 Then Exit Sub
borrar_todo
inicializa_todo
Label40.Visible = False
Frame5.Visible = False
Label17_Click
Frame7.Visible = True
pedido = ""
Label41_Click
pedido.SetFocus
End Sub
Function busca_codigo()
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("clientes")
mytablex.Index = "codigo"
mytablex.Seek "=", "" & Codigo
If Not mytablex.NoMatch Then
   busca_codigo = 1
   If Len(nombre) = 0 Then
      nombre = "" & mytablex.Fields("nombre")
   End If
End If
'------------------------------------- ------------
mytablex.Close

End Function

Private Sub observa_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   nombre.SetFocus
   Exit Sub
End If
End Sub

Private Sub pedido_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(pedido) = 0 Then
   pedido.SetFocus
   Exit Sub
End If
found = existe_pocket()
If found = 0 Then
   MsgBox "No existe Pedido Generado en la Entrada", 48, "Aviso"
   pedido = ""
   pedido.SetFocus
   Exit Sub
End If
xpedido = pedido
Frame7.Visible = False
borrar_todo
found = carga_proforma()
If found = 0 Then
   'MsgBox "No hay Datos", 48, "Aviso"
End If
SQL_pedido
found = suma_detalle()
found = ir_ultimo()
'If found = 0 Then
'   Data2.Refresh
'End If
Frame5.Visible = False
DBGrid2.SetFocus
End Sub
Function carga_proforma()
Dim i As Integer
Dim found As Integer
Dim mytablex As Table
On Error GoTo cmd67112_err
Set mytablex = mydbxglo.OpenTable("dproform")
mytablex.Index = "Tdetalle"
mytablex.Seek "=", xlocal, xtipo, xserie, pedido
If Not mytablex.NoMatch Then
Do
    If mytablex.EOF Then Exit Do
    If "" & mytablex.Fields("local") = xlocal And "" & mytablex.Fields("tipo") = xtipo And "" & mytablex.Fields("serie") = xserie And "" & mytablex.Fields("numero") = pedido Then
       Data2.Recordset.AddNew
       For i = 0 To mytablex.Fields.Count - 1
           Data2.Recordset.Fields(i) = mytablex.Fields(i)
       Next i
       Data2.Recordset.Update
       carga_proforma = 1
       Else: Exit Do
    End If
mytablex.MoveNext
Loop
End If
mytablex.Close
Exit Function
cmd67112_err:
mytablex.Close
MsgBox "x" & error$
Exit Function
End Function




Private Sub tipo_Change()

End Sub

Private Sub vendedor_KeyPress(KeyAscii As Integer)

End Sub

Private Sub vendedor_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   observa.SetFocus
   Exit Sub
End If
End Sub
Sub carga_inicial()
'Dim mytablex As Table
'Set mytablex = mydbxglo.OpenTable("vendedor")
'vendedor.Clear
'Do
'If mytablex.EOF Then Exit Do
'vendedor.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("NOMBRE")
'mytablex.MoveNext
'Loop
'vendedor.ListIndex = 0
'------------------------------------- ------------
'mytablex.Close

End Sub

Private Sub xcant_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
Label42_Click
End Sub

Private Sub xcodigo_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
inicializa_xproducto
found = busca_xproducto()
If found = 0 Then
   MsgBox "NO existe Producto", 48, "Aviso"
   xcodigo = ""
   xcodigo.SetFocus
   Exit Sub
End If
xcodigo.SetFocus

End Sub
Sub inicializa_xproducto()
xnombre = ""
xunidad = ""
xfactor = ""
xpventa = ""
xstock = ""
End Sub
Function existe_pocket()
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("ppocket")
mytablex.Index = "ppocket"
mytablex.Seek "=", pedido
If Not mytablex.NoMatch Then
   existe_pocket = 1
End If
mytablex.Close

End Function
Function existe_pedido()
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("cproform")
mytablex.Index = "tfactura"
mytablex.Seek "=", xlocal, xserie, xtipo, pedido
If Not mytablex.NoMatch Then
   existe_pedido = 1
End If
mytablex.Close

End Function
Function busca_xproducto()
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("producto")
mytablex.Index = "producto"
mytablex.Seek "=", xcodigo
If Not mytablex.NoMatch Then
   xnombre = "" & mytablex.Fields("descripcio")
   xunidad = "" & mytablex.Fields("unidad1")
   xfactor = "" & mytablex.Fields("factor1")
   xpventa = "" & mytablex.Fields("pventa1")
   busca_xproducto = 1
End If
mytablex.Close
End Function

