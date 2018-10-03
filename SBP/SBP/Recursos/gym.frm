VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form pocket 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Orion-PedidoPocket"
   ClientHeight    =   3285
   ClientLeft      =   150
   ClientTop       =   120
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Teclado"
      Height          =   1695
      Left            =   0
      TabIndex        =   82
      Top             =   1560
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "*"
         Height          =   255
         Index           =   41
         Left            =   1920
         TabIndex        =   125
         Top             =   480
         Width           =   375
      End
      Begin VB.Label xteclado 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   124
         Top             =   1680
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CLOSE"
         Height          =   255
         Index           =   40
         Left            =   2280
         TabIndex        =   123
         Top             =   1440
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
         TabIndex        =   122
         Top             =   480
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
         TabIndex        =   121
         Top             =   1440
         Width           =   735
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
         TabIndex        =   120
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Z"
         Height          =   255
         Index           =   36
         Left            =   480
         TabIndex        =   119
         Top             =   1440
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
         TabIndex        =   118
         Top             =   1440
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
         TabIndex        =   117
         Top             =   1200
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
         TabIndex        =   116
         Top             =   1200
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
         TabIndex        =   115
         Top             =   1200
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
         TabIndex        =   114
         Top             =   1200
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
         TabIndex        =   113
         Top             =   1200
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
         TabIndex        =   112
         Top             =   1200
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
         TabIndex        =   111
         Top             =   1200
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
         TabIndex        =   110
         Top             =   1200
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
         TabIndex        =   109
         Top             =   960
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
         TabIndex        =   108
         Top             =   960
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
         TabIndex        =   107
         Top             =   960
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
         TabIndex        =   106
         Top             =   960
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
         TabIndex        =   105
         Top             =   960
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
         TabIndex        =   104
         Top             =   960
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
         TabIndex        =   103
         Top             =   960
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
         TabIndex        =   102
         Top             =   960
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
         TabIndex        =   101
         Top             =   720
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
         Top             =   720
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
         TabIndex        =   99
         Top             =   720
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
         TabIndex        =   98
         Top             =   720
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
         TabIndex        =   97
         Top             =   720
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
         TabIndex        =   96
         Top             =   720
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
         TabIndex        =   95
         Top             =   720
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
         TabIndex        =   94
         Top             =   720
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
         TabIndex        =   93
         Top             =   480
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
         TabIndex        =   92
         Top             =   480
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
         TabIndex        =   91
         Top             =   480
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
         TabIndex        =   90
         Top             =   480
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
         TabIndex        =   89
         Top             =   480
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
         TabIndex        =   88
         Top             =   240
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
         TabIndex        =   87
         Top             =   240
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
         TabIndex        =   86
         Top             =   240
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
         TabIndex        =   85
         Top             =   240
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
         TabIndex        =   84
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   83
         Top             =   240
         Width           =   375
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
      Left            =   4800
      TabIndex        =   62
      Top             =   1920
      Visible         =   0   'False
      Width           =   3255
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
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
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
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   480
         Width           =   1815
      End
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
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox tipo 
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
         MaxLength       =   2
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox serie 
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
         MaxLength       =   3
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox numero 
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
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1215
      End
      Begin VB.ComboBox vendedor 
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
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   960
         Width           =   1815
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
         TabIndex        =   81
         Top             =   240
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
         TabIndex        =   80
         Top             =   480
         Width           =   735
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
         TabIndex        =   79
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFF00&
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
         Left            =   120
         TabIndex        =   78
         Top             =   960
         Width           =   735
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
         TabIndex        =   77
         Top             =   1800
         Width           =   495
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
         TabIndex        =   76
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
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
         TabIndex        =   75
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie"
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
         TabIndex        =   74
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero"
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
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label36 
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
         Left            =   2640
         TabIndex        =   72
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label37 
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
         Left            =   2640
         TabIndex        =   71
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label38 
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
         Left            =   2640
         TabIndex        =   70
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FF0000&
      Caption         =   "Lista Precios y Saldos "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3255
      Left            =   5160
      TabIndex        =   57
      Top             =   1080
      Visible         =   0   'False
      Width           =   3255
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   360
         Width           =   1215
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
         Height          =   495
         Left            =   1440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "gym.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Borrar registro"
         Top             =   360
         Width           =   495
      End
      Begin MSDBGrid.DBGrid DBGrid4 
         Height          =   2175
         Left            =   120
         OleObjectBlob   =   "gym.frx":1212
         TabIndex        =   60
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label tproducto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   420
         Left            =   2040
         TabIndex        =   61
         Top             =   360
         Width           =   135
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFF00&
      Caption         =   "Productos"
      Height          =   2535
      Left            =   4800
      TabIndex        =   39
      Top             =   960
      Visible         =   0   'False
      Width           =   3255
      Begin VB.TextBox xpventa 
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
         MaxLength       =   10
         TabIndex        =   126
         TabStop         =   0   'False
         Top             =   1560
         Width           =   975
      End
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
         MaxLength       =   15
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox xcantidad 
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
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label tmpventa 
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
         TabIndex        =   127
         Top             =   2040
         Visible         =   0   'False
         Width           =   735
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
         TabIndex        =   56
         Top             =   1200
         Width           =   615
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
         TabIndex        =   55
         Top             =   360
         Width           =   615
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
         TabIndex        =   54
         Top             =   1080
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
         TabIndex        =   53
         Top             =   600
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
         TabIndex        =   52
         Top             =   1320
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
         TabIndex        =   51
         Top             =   1560
         Width           =   735
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
         Left            =   1920
         TabIndex        =   50
         Top             =   1800
         Visible         =   0   'False
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
         TabIndex        =   49
         Top             =   1200
         Width           =   615
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
         TabIndex        =   48
         Top             =   1080
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
         TabIndex        =   47
         Top             =   600
         Width           =   2295
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
         TabIndex        =   46
         Top             =   1320
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
         Left            =   1920
         TabIndex        =   45
         Top             =   1560
         Visible         =   0   'False
         Width           =   735
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
         TabIndex        =   44
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label41 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cant"
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
         TabIndex        =   43
         Top             =   1800
         Width           =   735
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
         TabIndex        =   42
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ticket Ingreso Nro."
      Height          =   1575
      Left            =   7920
      TabIndex        =   29
      Top             =   1680
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox Text2 
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
         MaxLength       =   3
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text1 
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
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label48 
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
         Left            =   840
         TabIndex        =   35
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label47 
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
         Left            =   1440
         TabIndex        =   34
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label46 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie"
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
         TabIndex        =   33
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label44 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero"
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
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFF00&
      Caption         =   "Modifica"
      Height          =   1455
      Left            =   4440
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   3255
      Begin VB.TextBox xcant 
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
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label45 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cant"
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
         Left            =   960
         TabIndex        =   23
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label49 
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
         Left            =   1680
         TabIndex        =   22
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Graba"
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
         Left            =   2280
         TabIndex        =   21
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label35 
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
         Left            =   2280
         TabIndex        =   20
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Busqueda de Productos"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   3960
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   3255
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "gym.frx":1F2D
         Height          =   1335
         Left            =   120
         OleObjectBlob   =   "gym.frx":1F41
         TabIndex        =   28
         Top             =   840
         Width           =   3015
      End
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
         TabIndex        =   27
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
         TabIndex        =   26
         Top             =   240
         Width           =   1815
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
         TabIndex        =   25
         Top             =   2280
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
         TabIndex        =   17
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
         Left            =   2520
         TabIndex        =   15
         Top             =   360
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
         TabIndex        =   14
         Top             =   360
         Width           =   615
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Control Acceso"
      Height          =   3255
      Left            =   1800
      TabIndex        =   8
      Top             =   240
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
         MaxLength       =   2
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label vservidor 
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
         TabIndex        =   36
         Top             =   600
         Width           =   105
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
         TabIndex        =   11
         Top             =   2880
         Width           =   3015
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
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   1815
         Left            =   120
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Terminal"
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
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
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
      Bindings        =   "gym.frx":2914
      Height          =   2775
      Left            =   0
      OleObjectBlob   =   "gym.frx":2928
      TabIndex        =   1
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label Label40 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   0
      TabIndex        =   38
      Top             =   2760
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label43 
      BackColor       =   &H00FFFF00&
      Caption         =   "Lista"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   37
      Top             =   2760
      Width           =   495
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
      Left            =   0
      TabIndex        =   18
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4080
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Totaliza"
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
      TabIndex        =   12
      Top             =   3000
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
      TabIndex        =   7
      Top             =   3000
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
      TabIndex        =   6
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   3000
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
      Left            =   0
      TabIndex        =   4
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Total 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   2760
      Width           =   735
   End
End
Attribute VB_Name = "pocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type campo_precio
    unidad As String
    factor As String
    precio As String
    stock As String
End Type
Dim flag_especial As String
Dim campo_precios(12) As campo_precio

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
               'impresion_codbar buf
               
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

Private Sub clave_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(clave) = 0 Then
   clave.SetFocus
   Exit Sub
End If
clave = UCase(clave)
If Mid$(clave, 1, 1) <> "T" Then
   clave = ""
   MsgBox "No es Terminal ", 48, "Aviso"
   clave.SetFocus
   Exit Sub
End If
found = conectarpo()
If found = 0 Then
   MsgBox "Error de Conexion Sql Server ", 48, "Aviso"
   clave.SetFocus
   Exit Sub
End If

found = copiar_temporalp()
If found = 0 Then
   MsgBox "Terminal en Uso ", 48, "Aviso"
   clave.SetFocus
   Exit Sub
End If

 mytable11.Open "SELECT *  FROM parameca where caja='" & clave & "'", cn, adOpenKeyset, adLockOptimistic
 If mytable11.RecordCount = 0 Then
   clave = ""
   MsgBox "No existe Terminal", 48, "Aviso"
   mytable11.Close
   clave.SetFocus
   Exit Sub
End If
carga_inicial
gusuario = "_z" & clave
Frame3.Visible = False
SQL_pedido
found = suma_detalle()
DBGrid2.SetFocus
End Sub

Function busca_terminal()
Dim mytablex As New ADODB.Recordset
 mytablex.Open "SELECT *  FROM parameca where caja='" & clave & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount > 0 Then
      busca_terminal = 1
End If
mytablex.Close
End Function

Private Sub codigo_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(Codigo) > 0 Then
found = busca_codigo()
If found = 0 Then
End If
End If
nombre.SetFocus
End Sub

Private Sub codigo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   Frame2.Visible = False
   Exit Sub
End If
End Sub

Private Sub Combo2_Click()
If Len(tproducto) > 0 Then
DBGrid4.Refresh
carga_dbgrid4 tproducto, Combo2.Text
End If

End Sub

Private Sub Command8_Click()
Frame8.Visible = False
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim found As Integer
If KeyCode = 13 Then
   'adiciona_registro
   'suma_detalle
   'Frame1.Visible = False
   'Data2.Refresh
   'ir_ultimo
   'DBGrid2.SetFocus
   Label39_Click
End If

End Sub
Sub adiciona_registro()
Dim found As Integer
Data2.Recordset.AddNew
Data2.Recordset.Fields("producto") = "" & Data1.Recordset.Fields("producto")
Data2.Recordset.Fields("descripcio") = "" & Data1.Recordset.Fields("descripcio")
Data2.Recordset.Fields("unidad") = "" & Data1.Recordset.Fields("und")
Data2.Recordset.Fields("factor") = Val("" & Data1.Recordset.Fields("facT"))
Data2.Recordset.Fields("precio") = Val("" & Data1.Recordset.Fields("pvta"))
Data2.Recordset.Fields("igv") = Val("" & Data1.Recordset.Fields("igv"))
Data2.Recordset.Fields("total") = Val("" & Data1.Recordset.Fields("pvta"))
Data2.Recordset.Fields("cantidad") = 1
Data2.Recordset.Update
End Sub
Sub adiciona_registrox()
Dim sdx As Double
Data2.Recordset.AddNew
Data2.Recordset.Fields("producto") = "" & xcodigo
Data2.Recordset.Fields("descripcio") = "" & xnombre
Data2.Recordset.Fields("unidad") = "" & xunidad
Data2.Recordset.Fields("factor") = Val("" & xfactor)
Data2.Recordset.Fields("precio") = Val("" & xpventa)
Data2.Recordset.Fields("igv") = 18 'Val("" & xigv)
If Val(xcantidad) <= 0 Then
   xcantidad = "1"
End If
sdx = Val(xcantidad) * Val(xpventa)
Data2.Recordset.Fields("total") = Val(Format(sdx, "0.00"))
Data2.Recordset.Fields("cantidad") = Val(xcantidad)
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
            Data2.Refresh
            
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

Private Sub DBGrid4_DblClick()
DBGrid4_KeyDown 13, 0
End Sub
Function sumar_detalle()
Dim sdx As Double
Total = ""
Data2.Refresh
sdx = 0
Do
If Data2.Recordset.EOF Then Exit Do
sdx = sdx + Val("" & Data2.Recordset.Fields("total"))
Data2.Recordset.MoveNext
Loop
Total = Format(sdx, "0.00")
End Function

Private Sub DBGrid4_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sdx As Double
Dim found As Integer
Dim xpreciox As Double
If KeyCode = 27 Then
   Frame8.Visible = False
   found = sumar_detalle()
   DBGrid2.Col = 0
   DBGrid2.Row = DBGrid2.VisibleRows - 1
   DBGrid2.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
   If Len("" & DBGrid4.Columns(1)) = 0 Or Len("" & DBGrid4.Columns(0)) = 0 Then
      DBGrid4.SetFocus
      Exit Sub
   End If
   If Frame4.Visible = True Then
      xpreciox = 0
      xpreciox = Val("" & DBGrid4.Columns(2))
      'DBGrid2.Columns(51) = "" & (DBGrid4.Row + 1)
      tmpventa = xpreciox
      xunidad = "" & DBGrid4.Columns(0)
      xfactor = Val("" & DBGrid4.Columns(1))
      xpventa = xpreciox
      sdx = Val(xcantidad) * Val(xpventa)
      Frame8.Visible = False
      xcantidad.SetFocus
      Exit Sub
   
   End If
   
If Frame4.Visible = False Then
      Data2.Recordset.Edit
      xpreciox = 0
      xpreciox = Val("" & DBGrid4.Columns(2))
      'DBGrid2.Columns(51) = "" & (DBGrid4.Row + 1)
      Data2.Recordset.Fields("unidad") = "" & DBGrid4.Columns(0)
      Data2.Recordset.Fields("factor") = Val("" & DBGrid4.Columns(1))
      Data2.Recordset.Fields("precio") = xpreciox
      sdx = Val("" & Data2.Recordset.Fields("cantidad")) * Val("" & Data2.Recordset.Fields("precio"))
      Data2.Recordset.Fields("total") = sdx
      Data2.Recordset.Update
      found = sumar_detalle()
      Frame8.Visible = False
      DBGrid2.Col = 0
      DBGrid2.Row = DBGrid2.VisibleRows - 1
      DBGrid2.SetFocus
End If

End If
End Sub

Private Sub DBGrid4_UnboundReadData(ByVal RowBuf As MSDBGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim dr As Integer
Dim row_num As Integer
Dim r As Integer
Dim rows_returned As Integer
If ReadPriorRows Then
        dr = -1
    Else
        dr = 1
    End If
    If IsNull(StartLocation) Then
        If ReadPriorRows Then
           row_num = RowBuf.RowCount - 1
           'row_num = 9
        Else
           row_num = 0
        End If
    Else
        row_num = CLng(StartLocation) + dr
    End If
    rows_returned = 0
    For r = 0 To RowBuf.RowCount - 1
        If row_num < 0 Or row_num > 9 Then Exit For
        RowBuf.Value(r, 0) = campo_precios(row_num).unidad
        RowBuf.Value(r, 1) = campo_precios(row_num).factor
        RowBuf.Value(r, 2) = campo_precios(row_num).precio
        RowBuf.Value(r, 5) = campo_precios(row_num).stock
        RowBuf.Bookmark(r) = row_num
        row_num = row_num + dr
        rows_returned = rows_returned + 1
   Next r
   RowBuf.RowCount = rows_returned
End Sub


Private Sub Form_Load()
Dim found As Integer
Dim xxempresa As String
xxempresa = "Sistema Orion V5.0"
globaldir = App.path & "\001d\06"
'globaldir = "C:\ORION.V5\001D\06"
'globalpath = "C:\ORION.V5"
globalpath = "\ORION.V5"
'gusuario = "IDE"
'found = copiar_temporal()
'If found = 0 Then
'   MsgBox "Ya existe El uso ", 48, "Aviso"
'End If
carga_servidor
Set mydbxglo = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
condicion.Clear
condicion.AddItem "Producto.Producto"
condicion.AddItem "Producto.Descripcio"
condicion.ListIndex = 0
End Sub
Sub carga_dbgrid4(uproducto As String, xlistab As String)
Dim i As Integer
Dim paridad As String
Dim xfoto As String
Dim mytablex As New ADODB.Recordset
Dim mytabley As New ADODB.Recordset
Dim sw As Integer
Dim xbodega As String
Dim xsaldo As Double
Dim xbuf As String
Dim xcosto As Double
Dim xmargen As Double
Dim xcostou As Double
Dim xfactor As Double
Dim xxr As String
Dim xxi As String
Dim zbuf As String
Dim xpreciox As Double
Dim dmoneda As String
On Error GoTo cmd89111_err
xcostou = 0
For i = 0 To 9
    campo_precios(i).unidad = ""
    campo_precios(i).factor = ""
    campo_precios(i).precio = ""
    campo_precios(i).stock = ""
Next i
tproducto = uproducto
'MsgBox uproducto
xfactor = 1
xbodega = "" & mytable11.Fields("bodega")
xsaldo = 0
xcosto = 0
sw = 0
      If mytabley.State = 1 Then mytabley.Close
      mytabley.Open "SELECT * FROM almacen where local='" & "" & mytable11.Fields("local") & "' and producto='" & uproducto & "' and bodega='" & xbodega & "'", cn, adOpenStatic, adLockOptimistic
      If mytabley.RecordCount > 0 Then
         xsaldo = Val("" & mytabley.Fields("saldo"))
      End If
      mytabley.Close
'MsgBox "x"
'---buscamos los datos de productos
dmoneda = "S"
xfoto = ""
'descorto = ""

      If mytablex.State = 1 Then mytablex.Close
      mytablex.Open "SELECT * FROM producto where  producto='" & uproducto & "'", cn, adOpenStatic, adLockOptimistic
      If mytablex.RecordCount > 0 Then
         xcostou = 0
         If "" & mytable11.Fields("vecocaja") = "S" Then
            xcostou = Val("" & mytablex.Fields("costou"))
         End If
         xfactor = Val("" & mytablex.Fields("factor"))
         'descorto = "" & mytablex.Fields("presenta")
         dmoneda = "" & mytablex.Fields("monedav")
         xfoto = "" & mytablex.Fields("fotonombre")
      End If
      mytablex.Close
      'carga_foto xfoto
      'If Val(paridad) <= 0 Then
         paridad = "1"
      'End If
      If mytablex.State = 1 Then mytablex.Close
      
      '-------------------------------------------
    

If flag_especial = "S" Then
zbuf = "SELECT * FROM precio1 where  producto='" & uproducto & "' and local='01' and codigo='" & Codigo & "'"
      mytablex.Open zbuf, cn, adOpenStatic, adLockOptimistic
      If mytablex.RecordCount > 0 Then
         GoTo amika7
      End If
      mytablex.Close
End If

zbuf = "SELECT * FROM precios where  producto='" & uproducto & "' and local='" & xlistab & "'"
      mytablex.Open zbuf, cn, adOpenStatic, adLockOptimistic
amika7:
      If mytablex.RecordCount > 0 Then
         xcosto = 0
         xpreciox = 0
         If Val("" & mytablex.Fields("factor1")) > 0 Then
            If "" & mytable11.Fields("moneda") = "S" Then 'si es soles
               xpreciox = Val("" & mytablex.Fields("pventa1"))
               If dmoneda = "D" Then
                  xpreciox = Val("" & mytablex.Fields("pventa1")) * Val(paridad)
               End If
            End If
            If "" & mytable11.Fields("moneda") = "D" Then 'si es soles
               xpreciox = Val("" & mytablex.Fields("pventa1"))
               If dmoneda = "S" Then
                  xpreciox = Val("" & mytablex.Fields("pventa1")) / Val(paridad)
               End If
            End If
           '------------------------------------------------------------
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor1"))
            campo_precios(0).unidad = "" & mytablex.Fields("unidad1")
            campo_precios(0).factor = Val("" & mytablex.Fields("factor1"))
            campo_precios(0).precio = "" & xpreciox
            xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor1")))
            campo_precios(0).stock = "" & xbuf
            xmargen = 0
            If xcosto > 0 Then
               xmargen = ((xpreciox - xcosto) * 100) / xcosto
            End If
         End If
   '---------
   xcosto = 0
   If Val("" & mytablex.Fields("factor2")) > 0 Then
      If "" & mytable11.Fields("moneda") = "S" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa2"))
      If dmoneda = "D" Then
         xpreciox = Val("" & mytablex.Fields("pventa2")) * Val(paridad)
      End If
   End If
   If "" & mytable11.Fields("moneda") = "D" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa2"))
      If dmoneda = "S" Then
         xpreciox = Val("" & mytablex.Fields("pventa2")) / Val(paridad)
      End If
   End If
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor2"))
   campo_precios(1).unidad = "" & mytablex.Fields("unidad2")
   campo_precios(1).factor = Val("" & mytablex.Fields("factor2"))
   campo_precios(1).precio = xpreciox
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor2")))
   campo_precios(1).stock = "" & xbuf
   xmargen = 0
   If xcosto > 0 Then
      xmargen = ((xpreciox - xcosto) * 100) / xcosto
   End If
   End If
   xcosto = 0
   If Val("" & mytablex.Fields("factor3")) > 0 Then
      If "" & mytable11.Fields("moneda") = "S" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa3"))
      If dmoneda = "D" Then
         xpreciox = Val("" & mytablex.Fields("pventa3")) * Val(paridad)
      End If
   End If
   If "" & mytable11.Fields("moneda") = "D" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa3"))
      If dmoneda = "S" Then
         xpreciox = Val("" & mytablex.Fields("pventa3")) / Val(paridad)
      End If
   End If

   
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor3"))
   campo_precios(2).unidad = "" & mytablex.Fields("unidad3")
   campo_precios(2).factor = Val("" & mytablex.Fields("factor3"))
   campo_precios(2).precio = xpreciox
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor3")))
   campo_precios(2).stock = "" & xbuf
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((xpreciox - xcosto) * 100) / xcosto
   End If
   End If
   xcosto = 0
   If Val("" & mytablex.Fields("factor4")) > 0 Then
      If "" & mytable11.Fields("moneda") = "S" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa4"))
      If dmoneda = "D" Then
         xpreciox = Val("" & mytablex.Fields("pventa4")) * Val(paridad)
      End If
   End If
   If "" & mytable11.Fields("moneda") = "D" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa4"))
      If dmoneda = "S" Then
         xpreciox = Val("" & mytablex.Fields("pventa4")) / Val(paridad)
      End If
   End If

   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor4"))
   campo_precios(3).unidad = "" & mytablex.Fields("unidad4")
   campo_precios(3).factor = Val("" & mytablex.Fields("factor4"))
   campo_precios(3).precio = xpreciox
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor4")))
   campo_precios(3).stock = "" & xbuf
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((xpreciox - xcosto) * 100) / xcosto
   End If
   End If
   xcosto = 0
   If Val("" & mytablex.Fields("factor5")) > 0 Then
      If "" & mytable11.Fields("moneda") = "S" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa5"))
      If dmoneda = "D" Then
         xpreciox = Val("" & mytablex.Fields("pventa5")) * Val(paridad)
      End If
   End If
   If "" & mytable11.Fields("moneda") = "D" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa5"))
      If dmoneda = "S" Then
         xpreciox = Val("" & mytablex.Fields("pventa5")) / Val(paridad)
      End If
   End If

   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor5"))
      campo_precios(4).unidad = "" & mytablex.Fields("unidad5")
   campo_precios(4).factor = Val("" & mytablex.Fields("factor5"))
   campo_precios(4).precio = xpreciox
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor5")))
   campo_precios(4).stock = "" & xbuf
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((xpreciox - xcosto) * 100) / xcosto
   End If
   End If
   xcosto = 0
   
   If Val("" & mytablex.Fields("factor6")) > 0 Then
   
      If "" & mytable11.Fields("moneda") = "S" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa6"))
      If dmoneda = "D" Then
         xpreciox = Val("" & mytablex.Fields("pventa6")) * Val(paridad)
      End If
   End If
   If "" & mytable11.Fields("moneda") = "D" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa6"))
      If dmoneda = "S" Then
         xpreciox = Val("" & mytablex.Fields("pventa6")) / Val(paridad)
      End If
   End If

   
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor6"))
   campo_precios(5).unidad = "" & mytablex.Fields("unidad6")
   campo_precios(5).factor = Val("" & mytablex.Fields("factor6"))
   campo_precios(5).precio = xpreciox
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor6")))
   campo_precios(5).stock = "" & xbuf
   xmargen = 0
   If xcosto > 0 Then
      xmargen = ((xpreciox - xcosto) * 100) / xcosto
   End If
   'SOLO PARA MAXIMO SE PONE PRECIO=0
   'If caja <> "08" Then
   '   campo_precios(5).precio = 0
   'End If
   End If
   'MsgBox "xx"
   xcosto = 0
   If Val("" & mytablex.Fields("factor7")) > 0 Then
      If "" & mytable11.Fields("moneda") = "S" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa7"))
      If dmoneda = "D" Then
         xpreciox = Val("" & mytablex.Fields("pventa7")) * Val(paridad)
      End If
   End If
   If "" & mytable11.Fields("moneda") = "D" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa7"))
      If dmoneda = "S" Then
         xpreciox = Val("" & mytablex.Fields("pventa7")) / Val(paridad)
      End If
   End If

   
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor7"))
   
   campo_precios(6).unidad = "" & mytablex.Fields("unidad7")
   campo_precios(6).factor = Val("" & mytablex.Fields("factor7"))
   campo_precios(6).precio = xpreciox
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor7")))
   campo_precios(6).stock = "" & xbuf
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((xpreciox - xcosto) * 100) / xcosto
         
   End If
   End If
   
   xcosto = 0
   If Val("" & mytablex.Fields("factor8")) > 0 Then
   
      If "" & mytable11.Fields("moneda") = "S" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa8"))
      If dmoneda = "D" Then
         xpreciox = Val("" & mytablex.Fields("pventa8")) * Val(paridad)
      End If
   End If
   If "" & mytable11.Fields("moneda") = "D" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa8"))
      If dmoneda = "S" Then
         xpreciox = Val("" & mytablex.Fields("pventa8")) / Val(paridad)
      End If
   End If

      xcosto = xcostou / xfactor
      xcosto = xcosto * Val("" & mytablex.Fields("factor8"))
   
   campo_precios(7).unidad = "" & mytablex.Fields("unidad8")
   campo_precios(7).factor = Val("" & mytablex.Fields("factor8"))
   campo_precios(7).precio = xpreciox
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor8")))
   campo_precios(7).stock = "" & xbuf
   xmargen = 0
   If xcosto > 0 Then
      xmargen = ((xpreciox - xcosto) * 100) / xcosto
   End If
   End If
   xcosto = 0
   If Val("" & mytablex.Fields("factor9")) > 0 Then
      If "" & mytable11.Fields("moneda") = "S" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa9"))
      If dmoneda = "D" Then
         xpreciox = Val("" & mytablex.Fields("pventa9")) * Val(paridad)
      End If
   End If
   If "" & mytable11.Fields("moneda") = "D" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa9"))
      If dmoneda = "S" Then
         xpreciox = Val("" & mytablex.Fields("pventa9")) / Val(paridad)
      End If
   End If

   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor9"))
   campo_precios(8).unidad = "" & mytablex.Fields("unidad9")
   campo_precios(8).factor = Val("" & mytablex.Fields("factor9"))
   campo_precios(8).precio = xpreciox
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor9")))
   campo_precios(8).stock = "" & xbuf
   xmargen = 0
   If xcosto > 0 Then
      xmargen = ((xpreciox - xcosto) * 100) / xcosto
   End If
   End If
   xcosto = 0
   If Val("" & mytablex.Fields("factor10")) > 0 Then
      If "" & mytable11.Fields("moneda") = "S" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa10"))
      If dmoneda = "D" Then
         xpreciox = Val("" & mytablex.Fields("pventa10")) * Val(paridad)
      End If
   End If
   If "" & mytable11.Fields("moneda") = "D" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa10"))
      If dmoneda = "S" Then
         xpreciox = Val("" & mytablex.Fields("pventa10")) / Val(paridad)
      End If
   End If
   
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor10"))
   campo_precios(9).unidad = "" & mytablex.Fields("unidad10")
   campo_precios(9).factor = Val("" & mytablex.Fields("factor10"))
   campo_precios(9).precio = xpreciox
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor10")))
   campo_precios(9).stock = "" & xbuf
   xmargen = 0
   If xcosto > 0 Then
      xmargen = ((xpreciox - xcosto) * 100) / xcosto
   End If
   End If
   'MsgBox "xx"
   'sql_saldo_locales uproducto
   'margenes
   sw = 1
End If
'MsgBox ""
'mytablex.Close
'mytablez.Close
DBGrid4.Refresh
'----ahora deb cargar tambien la foto del producto...

'Frame1.Enabled = False
Frame8.Visible = True
DBGrid4.Enabled = True
DBGrid4.SetFocus
Exit Sub
cmd89111_err:
MsgBox "Error en carga dbgrid4 " + error$, 48, "Aviso"
Exit Sub
End Sub

Public Function conectarpo()
Dim dbuser As String
Dim dbpassword As String
Dim dbname As String
Dim dbserver As String
On Error GoTo cmd1_error
 
 cn.CursorLocation = adUseClient
 cn.CommandTimeout = 1024
 'cn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=calipso;Data Source=(local)"
' cn.Open "Driver={SQL Server};Server=" & menup.vservidor & ";Database=orion;Uid=sa"
 cn.Open "Driver={SQL Server};Server=" & vservidor & ";Database=calipso ;Uid=sa "
 
 'cn.Open "Driver={SQL Server};Server=ventas\kali;Database=calipso;Uid=sa"
 'cn.Open "Driver={SQL Server};Server=(local);Database=calipso;Uid=sa"
 conectarpo = 1
 Exit Function
cmd1_error:
 MsgBox " " & error$, 48, "Aviso"
 Exit Function
 End Function

Sub carga_servidor()
Dim found As Integer
Dim buf As String
On Error GoTo cmd169999_err
    vservidor = ""
    buf = ""
    If Dir$(globalpath & "\server.txt") <> "" Then
       Close
       Open globalpath & "\server.txt" For Input As #1
       Input #1, buf
       Close #1
       vservidor = buf
    End If
    '------------------------------------------------
    Exit Sub
cmd169999_err:
   Close
   Exit Sub

End Sub

Function suma_detalle()
Dim found As Integer
Dim xtotal As Double
Data2.Refresh
xtotal = 0
Do
If Data2.Recordset.EOF Then Exit Do
xtotal = xtotal + Val("" & Data2.Recordset.Fields("total"))
Data2.Recordset.MoveNext
Loop
Total = Format(xtotal, "0.00")
End Function
Function copiar_temporalp()
On Error GoTo cmd23_err
FileCopy globaldir & "\tdetalle.dbf", globaldir & "\" & "_z" & clave & ".dbf"
FileCopy globaldir & "\tdetalle.cdx", globaldir & "\" & "_z" & clave & ".cdx"
copiar_temporalp = 1
Exit Function
cmd23_err:
Exit Function
End Function
Sub SQL_pedido()
               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldir
               Data2.RecordSource = "select * from " & gusuario
               Data2.Refresh

End Sub

Private Sub Label10_Click()
Dim sdx As Double
Dim found As Integer
Label40.Visible = True
tipo = "C" '& mytable11.Fields("")
serie = "PO" & clave '& mytable11.Fields("")
sdx = Val("" & mytable11.Fields("numerope")) + 1
numero = "" & sdx
found = suma_detalle()
Frame2.Visible = True
Codigo.SetFocus
End Sub

Private Sub Label11_Click()
Dim found As Integer
Label40.Visible = False
found = suma_detalle()
Frame1.Visible = False
Data2.Refresh
DBGrid2.SetFocus
End Sub

Private Sub Label12_Click()
consulta_productos
End Sub

Private Sub Label16_Click()
numero_KeyPress 13
End Sub

Private Sub Label17_Click()
Dim found  As Integer
Label40.Visible = False
found = suma_detalle()
Data2.Refresh
Frame2.Visible = False
Label40.Visible = False
DBGrid2.SetFocus
End Sub

Private Sub Label22_Click()
Dim found As Integer
Label40.Visible = False
found = suma_detalle()
Frame4.Visible = False
Data2.Refresh
DBGrid2.SetFocus

End Sub

Private Sub Label23_Click()
If xcodigo.Enabled = True Then
   xcodigo.SetFocus
   Exit Sub
End If

xcodigo_KeyPress 13
End Sub

Private Sub Label29_Click()
Dim found As Integer
Dim mytablex As New ADODB.Recordset
If xcodigo.Enabled = True Then
   xcodigo.SetFocus
   Exit Sub
End If
If Val(xcantidad) <= 0 Then
   xcantidad.SetFocus
   Exit Sub
End If
'found = busca_xproducto("" & xcodigo)
'If found = 0 Then
'   xcodigo.SetFocus
'   Exit Sub
'End If
If Len(xunidad) = 0 Then Exit Sub
If Val(xfactor) = 0 Then Exit Sub
If Val(xpventa) <= 0 Then Exit Sub

If Val(tmpventa) > Val(xpventa) Then
   xpventa = tmpventa
   Exit Sub
End If


   xcodigo.Enabled = True
   adiciona_registrox
   found = suma_detalle()
   Frame4.Visible = False
   Data2.Refresh
   ir_ultimo
   Label40.Visible = False
   DBGrid2.SetFocus

End Sub

Private Sub Label3_Click()
xcodigo.Enabled = True
Label40.Visible = True
Frame4.Visible = True
inicializa_xproducto
xcodigo = ""
xcodigo.SetFocus
End Sub

Private Sub Label30_Click()
xteclado = "Producto"
Frame5.Visible = True

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
               codigo_KeyPress 13
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
If xteclado = "xcantidad" Then
   Select Case Index
          Case 39 'emter
               Frame5.Visible = False
               If Frame4.Visible = False Then Exit Sub
               xcantidad_KeyPress 13
               Exit Sub
          Case 37 'delete
               xcantidad = ""
               Exit Sub
          Case 40 'close
               Frame5.Visible = False
               If Frame4.Visible = False Then Exit Sub
               xcantidad.SetFocus
               Exit Sub
   End Select
   xcantidad = xcantidad & Label31(Index)
   Exit Sub
End If


If xteclado = "Clave" Then
   Select Case Index
          Case 39 'emter
               Frame5.Visible = False
               If Frame3.Visible = False Then Exit Sub
               clave_KeyPress 13
               Exit Sub
          Case 37 'delete
               xcodigo = ""
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
               If Frame6.Visible = False Then Exit Sub
               xcant_KeyPress 13
               Exit Sub
          Case 37 'delete
               xcant = ""
               Exit Sub
          Case 40 'close
               Frame5.Visible = False
               If Frame6.Visible = False Then Exit Sub
               
               xcant.SetFocus
               Exit Sub
   End Select
   xcant = xcant & Label31(Index)
End If

End Sub

Private Sub Label33_Click()
xteclado = "Busca"
Frame5.Visible = True

End Sub

Private Sub Label34_Click()
On Error GoTo cmd123_err
xcant = "" & Data2.Recordset.Fields("cantidad")
Frame6.Visible = True
xcant.SetFocus
Exit Sub
cmd123_err:
Data2.Refresh
Exit Sub

End Sub

Private Sub Label35_Click()
xteclado = "Cant"
Frame5.Visible = True

End Sub

Private Sub Label36_Click()
xteclado = "Codigo"
Frame5.Visible = True

End Sub

Private Sub Label37_Click()
xteclado = "Nombre"
Frame5.Visible = True

End Sub

Private Sub Label38_Click()
xteclado = "Observa"
Frame5.Visible = True

End Sub

Private Sub Label39_Click()
Dim found As Integer
   adiciona_registro
   found = sumar_detalle()
   'suma_detalle
   Frame1.Visible = False
   Data2.Refresh
   ir_ultimo
   Label40.Visible = False
   DBGrid2.SetFocus

End Sub

Private Sub Label4_Click()
Label40.Visible = True
buffer = ""
Frame1.Visible = True
consulta_productos
dbgrid1.SetFocus
Label33_Click

End Sub
Sub consulta_productos()
Dim buf1 As String
buf1 = "select Producto.Descripcio,Precios.Pventa1 as Pvta,Producto.producto,Producto.Marca,Producto.Monedav as M,Producto.Familia,Precios.Unidad1 as Und,Precios.Factor1 as Fact,Producto.Igv  from producto  left join precios on producto.producto=precios.producto  where precios.local='" & "" & mytable11.Fields("listap") & "' and " & condicion & " like '" & buffer & "*'"

               Data1.Connect = "foxpro 2.5;"
               Data1.DatabaseName = globaldir
               Data1.RecordSource = buf1
               Data1.Refresh
               dbgrid1.Columns(0).Width = 1800
               dbgrid1.Columns(1).Width = 500


End Sub
Sub borrar_todo()
Dim found As Integer
On Error GoTo cmd90_err
found = ir_inicio()
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
xteclado = "xcantidad"
Frame5.Visible = True

End Sub
Sub carga_combo2(buf As String)
Dim i As Integer
Combo2.Clear
      Combo2.AddItem buf
      For i = 1 To 9
         If buf <> Format(i, "00") Then
            Combo2.AddItem Format(i, "00")
         End If
      Next i
      Combo2.ListIndex = 0
End Sub


Private Sub Label42_Click()
Dim sdx As Double
Dim found As Integer
On Error GoTo cmd8911_err
If Not IsNumeric(xcant) Then
   Exit Sub
End If
Data2.Recordset.Edit
Data2.Recordset.Fields("cantidad") = Val(xcant)
sdx = Val(xcant) * Val("" & Data2.Recordset.Fields("precio"))
Data2.Recordset.Fields("total") = Val(Format(sdx, "0.00"))
Data2.Recordset.Update
Label40.Visible = False
found = sumar_detalle()
Label49_Click
Exit Sub
cmd8911_err:
MsgBox "No se puede Grabar", 48, "Aviso"
Exit Sub
End Sub


Private Sub Label43_Click()
On Error GoTo cmd89_err
Dim xproducto As String
If Len("" & Data2.Recordset.Fields("producto")) > 0 Then
      xproducto = "" & Data2.Recordset.Fields("producto")
      tproducto = ""
      carga_combo2 "" & mytable11.Fields("listap")
      carga_dbgrid4 "" & Data2.Recordset.Fields("producto"), "" & mytable11.Fields("listap")
      Exit Sub
End If
Exit Sub
cmd89_err:
Exit Sub
End Sub

Private Sub Label49_Click()
Label40.Visible = False
Frame6.Visible = False
Data2.Refresh
DBGrid2.SetFocus
End Sub

Private Sub Label5_Click()
Dim found As Integer
borra_linea
found = suma_detalle()
Data2.Refresh
End Sub

Private Sub Label6_Click()
Dim found As Integer
borrar_todo
found = suma_detalle()

End Sub

Private Sub Label7_Click()
xteclado = "Clave"
Frame5.Visible = True

End Sub

Private Sub Label8_Click()
clave_KeyPress 13
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

Private Sub numero_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
found = grabar()
If found = 0 Then Exit Sub
borrar_todo
inicializa_todo
Label40.Visible = False
Label17_Click
End Sub
Function busca_codigo()
Dim mytablex As New ADODB.Recordset
 mytablex.Open "SELECT *  FROM clientes where codigo='" & Codigo & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount > 0 Then
      busca_codigo = 1
      If Len(nombre) = 0 Then
      nombre = "" & mytablex.Fields("nombre")
      End If
End If
'------------------------------------- ------------
mytablex.Close
End Function
Function busca_vendedor()
Dim mytablex As New ADODB.Recordset
mytablex.Open "SELECT *  FROM vendedor where codigo='" & vendedor & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount > 0 Then
    busca_vendedor = 1
End If
'------------------------------------- ------------
mytablex.Close

End Function
Function grabar()
Dim found As Integer
found = valida()
If found = 0 Then
   MsgBox "Campos Invalidos", 48, "Aviso"
   Exit Function
End If
found = numero_libre()
found = adiciona_cliente()
found = adiciona_cabeza()
found = adiciona_detalle()
grabar = 1
End Function
Function adiciona_cliente()
Dim mytablex As New ADODB.Recordset

mytablex.Open "SELECT *  FROM clientes where codigo='" & Codigo & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount > 0 Then
   mytablex.Fields("codigo") = "" & Codigo
   mytablex.Fields("nombre") = "" & nombre
   mytablex.Update
Else
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
If Len(nombre) = 0 Then
   nombre.SetFocus
   Exit Function
End If
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
Dim found As Integer
Dim sdx As Double
Dim mytablex As New ADODB.Recordset
sdx = Val("" & numero)
RECIBE:
mytablex.Open "SELECT *  FROM cproform where local='" & mytable11.Fields("local") & "' and tipo='" & tipo & "' and serie='" & serie & "' and numero='" & numero & "'", cn, adOpenKeyset, adLockOptimistic
If mytablex.RecordCount > 0 Then
   sdx = sdx + 1
   numero = "" & sdx
   mytablex.Close
   GoTo RECIBE
End If
mytablex.Close

End Function
Function adiciona_cabeza()
Dim mytablex As New ADODB.Recordset
mytablex.Open "SELECT *  FROM cproform where local='" & mytable11.Fields("local") & "' and tipo='" & tipo & "' and serie='" & serie & "' and numero='" & numero & "'", cn, adOpenKeyset, adLockOptimistic
If mytablex.RecordCount = 0 Then
 
   mytablex.AddNew
   mytablex.Fields("local") = "" & mytable11.Fields("local")
   mytablex.Fields("tipo") = "" & tipo
   mytablex.Fields("serie") = "" & serie
   mytablex.Fields("numero") = "" & numero
   mytablex.Fields("codigo") = "" & Codigo
   mytablex.Fields("nombre") = "" & nombre
   mytablex.Fields("moneda") = "" & mytable11.Fields("moneda")
   mytablex.Fields("bodega") = "" & mytable11.Fields("bodega")
   mytablex.Fields("estado") = "2"
   mytablex.Fields("usuario") = extra_loquesea("" & vendedor)
   mytablex.Fields("total") = Val("" & Total)
   mytablex.Fields("CAJA") = "" & clave
   mytablex.Fields("tipoclie") = "C"
   mytablex.Fields("ACU") = "G"
   mytablex.Fields("SERVICIO") = "*"
   mytablex.Fields("vendedor") = extra_loquesea("" & vendedor)
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
Dim rs
Dim found As Integer
Dim mytablex As New ADODB.Recordset
cn.Execute ("delete   FROM dproform where local='" & mytable11.Fields("local") & "' and tipo='" & tipo & "' and serie='" & serie & "' and numero='" & numero & "'")
mytablex.Open "SELECT *  FROM dproform where local='" & mytable11.Fields("local") & "' and tipo='" & tipo & "' and serie='" & serie & "' and numero='" & numero & "'", cn, adOpenKeyset, adLockOptimistic
If mytablex.RecordCount > 0 Then
   mytablex.Close
   Exit Function
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
Sub graba_detalle(mytablex As ADODB.Recordset)
   mytablex.Fields("local") = "" & mytable11.Fields("local")
   mytablex.Fields("tipo") = "" & tipo
   mytablex.Fields("serie") = "" & serie
   mytablex.Fields("numero") = "" & numero
   mytablex.Fields("codigo") = "" & Codigo
   'mytablex.Fields("nombre") = "" & nombre
   mytablex.Fields("vendedor") = extra_loquesea("" & vendedor)
   mytablex.Fields("fecha") = Format(Now, "dd/mm/yyyy")
   mytablex.Fields("hora") = Format(Now, "hh:mm:ss")
   mytablex.Fields("moneda") = "" & mytable11.Fields("moneda")
   mytablex.Fields("bodega") = "" & mytable11.Fields("bodega")
   mytablex.Fields("estado") = "2"
   mytablex.Fields("usuario") = extra_loquesea("" & vendedor)
End Sub
Sub inicializa_todo()
Codigo = ""
nombre = ""
vendedor.ListIndex = 0
observa = ""
End Sub

Private Sub numero_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   vendedor.SetFocus
   Exit Sub
End If

End Sub

Private Sub observa_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
vendedor.SetFocus

End Sub

Private Sub observa_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   nombre.SetFocus
   Exit Sub
End If
End Sub

Private Sub vendedor_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
numero.SetFocus
End Sub

Private Sub vendedor_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   observa.SetFocus
   Exit Sub
End If
End Sub
Sub carga_inicial()
Dim mytablex As New ADODB.Recordset
vendedor.Clear
mytablex.Open "SELECT *  FROM vendedor ", cn, adOpenKeyset, adLockOptimistic
Do
If mytablex.EOF Then Exit Do
vendedor.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("NOMBRE")
mytablex.MoveNext
Loop
vendedor.ListIndex = 0
mytablex.Close

End Sub

Private Sub xcant_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
Label42_Click
End Sub

Private Sub xcantidad_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
Label29_Click

End Sub

Private Sub xcodigo_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(xcodigo) = 0 Then
   xcodigo.SetFocus
   Exit Sub
End If
inicializa_xproducto
found = busca_xproducto("" & xcodigo)
If found = 0 Then
   MsgBox "NO existe Producto", 48, "Aviso"
   xcodigo = ""
   xcodigo.SetFocus
   Exit Sub
End If
      xcodigo.Enabled = False
      carga_combo2 "" & mytable11.Fields("listap")
      carga_dbgrid4 "" & xcodigo, "" & mytable11.Fields("listap")
Exit Sub
cmd89_err:
Exit Sub

'xcantidad = "1"
'xcantidad.SetFocus

End Sub
Sub inicializa_xproducto()
xnombre = ""
xunidad = ""
xfactor = ""
xpventa = ""
xcantidad = ""
xstock = ""
End Sub
Function busca_xproducto(buf As String)
Dim mytablex As New ADODB.Recordset
Dim mytabley As New ADODB.Recordset
Dim found As Integer
Dim buf1 As String
Dim i As Integer
Dim ssw As Integer
Dim sw As Integer
i = 0

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
   xnombre = "" & mytablex.Fields("descripcio")
   xcodigo = "" & mytablex.Fields("producto")
   mytablex.Close
   mytablex.Open "select * from precios  where producto='" & xcodigo & "' and local='" & mytable11.Fields("listaP") & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount = 0 Then
      mytablex.Close
      Exit Function
   End If
   xunidad = "" & mytablex.Fields("unidad1")
   xfactor = "" & mytablex.Fields("factor1")
   xpventa = "" & mytablex.Fields("pventa1")
   tmpventa = "" & mytablex.Fields("pventa1")
   mytablex.Close
   busca_xproducto = 1
End Function
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

