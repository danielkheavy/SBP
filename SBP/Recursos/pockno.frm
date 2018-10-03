VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form pockno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Pedidos Normal"
   ClientHeight    =   11460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11460
   ScaleWidth      =   14445
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Teclado"
      Height          =   1695
      Left            =   8400
      TabIndex        =   119
      Top             =   1440
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
         TabIndex        =   162
         Top             =   480
         Width           =   375
      End
      Begin VB.Label xteclado 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   161
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
         TabIndex        =   160
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
         TabIndex        =   159
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
         TabIndex        =   158
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
         TabIndex        =   157
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
         TabIndex        =   156
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
         TabIndex        =   155
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
         TabIndex        =   154
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
         TabIndex        =   153
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
         TabIndex        =   152
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
         TabIndex        =   151
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
         TabIndex        =   150
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
         TabIndex        =   149
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
         TabIndex        =   148
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
         TabIndex        =   147
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
         TabIndex        =   146
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
         TabIndex        =   145
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
         TabIndex        =   144
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
         TabIndex        =   143
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
         TabIndex        =   142
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
         TabIndex        =   141
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
         TabIndex        =   140
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
         TabIndex        =   139
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
         TabIndex        =   138
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
         TabIndex        =   137
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
         TabIndex        =   136
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
         TabIndex        =   135
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
         TabIndex        =   134
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
         TabIndex        =   133
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
         TabIndex        =   132
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
         TabIndex        =   131
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
         TabIndex        =   130
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
         TabIndex        =   129
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
         TabIndex        =   128
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
         TabIndex        =   127
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
         TabIndex        =   126
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
         TabIndex        =   125
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
         TabIndex        =   124
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
         TabIndex        =   123
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
         TabIndex        =   122
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
         TabIndex        =   121
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
         TabIndex        =   120
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Ingreso de Lineas"
      Height          =   3255
      Left            =   8400
      TabIndex        =   78
      Top             =   3240
      Visible         =   0   'False
      Width           =   3255
      Begin VB.TextBox t1 
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox t2 
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t3 
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox t4 
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox t5 
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t6 
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox t7 
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox t8 
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox t9 
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox t10 
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox t11 
         Height          =   285
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox t12 
         Height          =   285
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t13 
         Height          =   285
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox t14 
         Height          =   285
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox t15 
         Height          =   285
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t16 
         Height          =   285
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   1800
         Width           =   735
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
         Height          =   495
         Left            =   2520
         MaskColor       =   &H00E0E0E0&
         Picture         =   "pockno.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Grabar registro"
         Top             =   2160
         Width           =   495
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
         Height          =   495
         Left            =   1920
         MaskColor       =   &H00E0E0E0&
         Picture         =   "pockno.frx":1212
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Borrar registro"
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Linea"
         Height          =   375
         Left            =   120
         TabIndex        =   116
         Top             =   240
         Width           =   495
      End
      Begin VB.Label nlinea 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   600
         TabIndex        =   115
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tallas"
         Height          =   255
         Left            =   120
         TabIndex        =   114
         Top             =   600
         Width           =   495
      End
      Begin VB.Label nt1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   600
         TabIndex        =   113
         Top             =   600
         Width           =   495
      End
      Begin VB.Label nt2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   600
         TabIndex        =   112
         Top             =   840
         Width           =   495
      End
      Begin VB.Label nt3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   600
         TabIndex        =   111
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label nt4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   600
         TabIndex        =   110
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label nt5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   600
         TabIndex        =   109
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label nt6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   600
         TabIndex        =   108
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label nt7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   600
         TabIndex        =   107
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label nt8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   600
         TabIndex        =   106
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label nt9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   600
         TabIndex        =   105
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label nt10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   600
         TabIndex        =   104
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label nt11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   103
         Top             =   600
         Width           =   735
      End
      Begin VB.Label nt12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   102
         Top             =   840
         Width           =   735
      End
      Begin VB.Label nt13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   101
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label nt14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   100
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label nt15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   99
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   98
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label linea 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2520
         TabIndex        =   97
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFF00&
      Caption         =   "Modifica"
      Height          =   1455
      Left            =   8400
      TabIndex        =   41
      Top             =   6600
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
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
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
         TabIndex        =   45
         Top             =   1080
         Width           =   615
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
         TabIndex        =   44
         Top             =   1080
         Width           =   615
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
         TabIndex        =   43
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFF00&
      Caption         =   "Productos"
      Height          =   2535
      Left            =   8400
      TabIndex        =   23
      Top             =   8160
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
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label xlinea 
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
         TabIndex        =   118
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Linea"
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
         TabIndex        =   117
         Top             =   2040
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         Left            =   1560
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
         Top             =   1320
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
         TabIndex        =   29
         Top             =   1560
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
         Left            =   2280
         TabIndex        =   28
         Top             =   1800
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
         TabIndex        =   27
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
         TabIndex        =   26
         Top             =   1800
         Width           =   735
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   1140
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
      Height          =   3135
      Left            =   4920
      TabIndex        =   14
      Top             =   7440
      Visible         =   0   'False
      Width           =   3255
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
         TabIndex        =   17
         Top             =   240
         Width           =   1815
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
         TabIndex        =   16
         Top             =   480
         Width           =   1815
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "pockno.frx":2424
         Height          =   2175
         Left            =   120
         OleObjectBlob   =   "pockno.frx":2438
         TabIndex        =   15
         Top             =   840
         Width           =   3015
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   120
         Width           =   615
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
         TabIndex        =   18
         Top             =   600
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
      Height          =   300
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   3720
      TabIndex        =   2
      Top             =   2040
      Width           =   3255
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   29
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   28
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   27
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   26
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   25
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   24
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   23
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   22
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   21
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   720
         Width           =   495
      End
      Begin VB.ComboBox vendedor 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox ticket 
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
         TabIndex        =   0
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "GeneraTicket"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   150
         Left            =   120
         TabIndex        =   163
         Top             =   3000
         Width           =   825
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Refresca"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   77
         Top             =   2520
         Width           =   975
      End
      Begin VB.Image Image5 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2520
         Picture         =   "pockno.frx":2E03
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   480
      End
      Begin VB.Image Image6 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2040
         Picture         =   "pockno.frx":4DA9
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   480
      End
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ticket Nro"
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
         TabIndex        =   4
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Entrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   2280
         TabIndex        =   3
         Top             =   120
         Width           =   735
      End
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "pockno.frx":697B
      Height          =   2775
      Left            =   0
      OleObjectBlob   =   "pockno.frx":698F
      TabIndex        =   1
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      Height          =   255
      Left            =   1920
      TabIndex        =   46
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label40 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   1680
      TabIndex        =   22
      Top             =   5520
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Total 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   2760
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
      Left            =   720
      TabIndex        =   10
      Top             =   5280
      Visible         =   0   'False
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
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   3000
      Width           =   975
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
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   3000
      Width           =   735
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
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   3000
      Width           =   735
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
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      DataField       =   "linea"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   30
      TabIndex        =   5
      Top             =   2760
      Width           =   375
   End
End
Attribute VB_Name = "pockno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ulvendedor As String
Dim mfamcod(15000) As String
Dim mfampag As Integer
Dim mfamtop As Integer
Option Explicit

Sub SQL_pedido()
Dim buf As String
On Error GoTo cmd1_err
buf = "select * from " & "_po" & ticket & " where local='01' and tipo='PO' and serie='001' and numero='" & ticket & "'"
               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldir
               Data2.RecordSource = buf
               Data2.Refresh
               Exit Sub
cmd1_err:
MsgBox "Aviso en slq_pedido " + error$, 48, "Aviso"
Exit Sub
End Sub


Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then
   'MsgBox ""
   'consulta_productos
   Exit Sub
End If
If KeyAscii = 13 Then
   consulta_productos
   DBGrid1.SetFocus
   Else
End If

End Sub

Private Sub buffer_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> 13 And KeyCode <> 27 Then
   consulta_productos
End If

End Sub

Private Sub Command1_Click(Index As Integer)
ticket = Command1(Index).Caption


End Sub

Private Sub Command2_Click()
Dim sdx As Double
Dim found As Integer
Data2.Recordset.Edit
Data2.Recordset.Fields("t1") = Val(t1)
Data2.Recordset.Fields("t2") = Val(t2)
Data2.Recordset.Fields("t3") = Val(t3)
Data2.Recordset.Fields("t4") = Val(t4)
Data2.Recordset.Fields("t5") = Val(t5)
Data2.Recordset.Fields("t6") = Val(t6)
Data2.Recordset.Fields("t7") = Val(t7)
Data2.Recordset.Fields("t8") = Val(t8)
Data2.Recordset.Fields("t9") = Val(t9)
Data2.Recordset.Fields("t10") = Val(t10)
Data2.Recordset.Fields("t11") = Val(t11)
Data2.Recordset.Fields("t12") = Val(t12)
Data2.Recordset.Fields("t13") = Val(t13)
Data2.Recordset.Fields("t14") = Val(t14)
Data2.Recordset.Fields("t15") = Val(t15)

sdx = Val(t1) + Val(t2) + Val(t3) + Val(t4) + Val(t5) + Val(t6) + Val(t7) + Val(t8) + Val(t9) + Val(t10) + Val(t11) + Val(t12) + Val(t13) + Val(t14) + Val(t15) + Val(t16)
Data2.Recordset.Fields("cantidad") = sdx
sdx = Val(sdx) * Val("" & Data2.Recordset.Fields("precio"))
Data2.Recordset.Fields("total") = Val(Format(sdx, "0.00"))
Data2.Recordset.Update
Label40.Visible = False
suma_detalle
Frame2.Visible = False
Label49_Click

End Sub

Private Sub Command3_Click()
Dim found As Integer
Label40.Visible = False
Frame6.Visible = False
Data2.Refresh
DBGrid2.SetFocus
Frame2.Visible = False
Frame5.Visible = False
found = ir_ultimo()
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim found As Integer
If KeyCode = 13 Then
   Label39_Click
End If

End Sub

Private Sub DBGrid2_AfterColUpdate(ByVal ColIndex As Integer)
Dim found As Integer
Select Case ColIndex
       Case 1, 3
            found = suma_detalle()
            found = ir_ultimo()
            
End Select
End Sub

Private Sub DBGrid2_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
Select Case ColIndex
       Case 0
            Cancel = True
            Exit Sub
       Case 3, 4, 5, 6, 7
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
            DBGrid2.Columns(3) = Val("" & DBGrid2.Columns(1)) * Val("" & DBGrid2.Columns(2))
       Case 2
            If Not IsNumeric(DBGrid2.Columns(2)) Then
               Cancel = True
               Exit Sub
            End If
            DBGrid2.Columns(3) = Val("" & DBGrid2.Columns(1)) * Val("" & DBGrid2.Columns(2))
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

Sub adiciona_registrox()
Dim sdx As Double
Data2.Recordset.AddNew
Data2.Recordset.Fields("local") = "01"
Data2.Recordset.Fields("tipo") = "PO"
Data2.Recordset.Fields("SERIE") = "001"
Data2.Recordset.Fields("numero") = ticket
Data2.Recordset.Fields("linea") = xlinea
Data2.Recordset.Fields("producto") = "" & xcodigo
Data2.Recordset.Fields("descripcio") = "" & xnombre
Data2.Recordset.Fields("unidad") = "" & xunidad
Data2.Recordset.Fields("factor") = Val("" & xfactor)
Data2.Recordset.Fields("precio") = Val("" & xpventa)
Data2.Recordset.Fields("igv") = 19 'Val("" & xigv)
If Val(xcantidad) <= 0 Then
   xcantidad = "1"
End If
sdx = Val(xcantidad) * Val(xpventa)
Data2.Recordset.Fields("total") = Val(Format(sdx, "0.00"))
Data2.Recordset.Fields("cantidad") = Val(xcantidad)
Data2.Recordset.Update
DBGrid1.Col = 1
End Sub

Private Sub Form_Load()
globaldir = App.Path & "\001d\06"
globaldir = "c:\orion.v5\001D\06"
Set mydbxglo = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
carga_inicial
carga_familia
condicion.Clear
condicion.AddItem "Producto.Producto"
condicion.AddItem "Producto.Descripcio"
condicion.ListIndex = 0
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

Function ir_ultimo()
On Error GoTo cmd90_err
Data2.Recordset.MoveLast
Exit Function
cmd90_err:
Exit Function
End Function

Function suma_detalle()
Dim found As Integer
Dim xtotal As Double
found = ir_inicio()
xtotal = 0
Do
If Data2.Recordset.EOF Then Exit Do
xtotal = xtotal + Val("" & Data2.Recordset.Fields("total"))
Data2.Recordset.MoveNext
Loop
Total = Format(xtotal, "0.00")
End Function

Sub adiciona_registro()
Dim found As Integer
Data2.Recordset.AddNew
Data2.Recordset.Fields("local") = "01"
Data2.Recordset.Fields("tipo") = "PO"
Data2.Recordset.Fields("SERIE") = "001"
Data2.Recordset.Fields("numero") = ticket
Data2.Recordset.Fields("linea") = "" & Data1.Recordset.Fields("linea")
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

Sub consulta_productos()
Dim buf1 As String
buf1 = "select Producto.Descripcio,Precios.Pventa1 as Pvta,Producto.producto,Producto.Marca,Producto.Monedav as M,Producto.Familia,Precios.Unidad1 as Und,Precios.Factor1 as Fact,Producto.Igv,Producto.Linea  from producto  left join precios on producto.producto=precios.producto  where precios.local='01' and " & condicion & " like '" & buffer & "*'"
'MsgBox buf1

               Data1.Connect = "foxpro 2.5;"
               Data1.DatabaseName = globaldir
               Data1.RecordSource = buf1
               Data1.Refresh
               DBGrid1.Columns(0).Width = 1800
               DBGrid1.Columns(1).Width = 500


End Sub

Sub carga_inicial()
Dim mytablex As Table
vendedor.Clear
Set mytablex = mydbxglo.OpenTable("vendedor")
vendedor.AddItem "*"
Do
If mytablex.EOF Then Exit Do
vendedor.AddItem "" & mytablex.Fields("Codigo") & "|" & mytablex.Fields("Nombre")
mytablex.MoveNext
Loop
vendedor.ListIndex = 0
mytablex.Close

End Sub

Private Sub Image5_Click()
menu_familia "SIG"
End Sub

Private Sub Image6_Click()
menu_familia "ANT"
End Sub

Private Sub Label1_Click()
carga_inicial

End Sub

Private Sub Label10_Click()
Dim found As Integer
If Val(Total) <= 0 Then
   MsgBox "Total=0", 48, "Aviso"
   Exit Sub
End If
If MsgBox("Desea Grabar", 1, "Aviso") <> 1 Then Exit Sub
found = adiciona_cabeza()
found = adiciona_detalle()
Frame3.Visible = True
ticket = ""
Label6_Click
vendedor.Clear
vendedor.AddItem ulvendedor
vendedor.ListIndex = 0
ticket.SetFocus


End Sub

Private Sub Label11_Click()
Dim found As Integer
Label40.Visible = False
found = suma_detalle()
Frame1.Visible = False
Data2.Refresh
DBGrid2.SetFocus
Frame5.Visible = False
Frame5.Visible = False

End Sub

Private Sub Label12_Click()
consulta_productos

End Sub

Private Sub Label13_Click()
Dim found As Integer
found = prepara_achivo()
Label2_Click

End Sub

Private Sub Label2_Click()
carga_familia

End Sub

Private Sub Label22_Click()
Dim found As Integer
Label40.Visible = False
found = suma_detalle()
Frame4.Visible = False
Data2.Refresh
DBGrid2.SetFocus
Frame5.Visible = False

End Sub

Private Sub Label23_Click()
xcodigo_KeyPress 13
End Sub

Private Sub Label29_Click()
Dim found As Integer
found = busca_xproducto()
If found = 0 Then
   xcodigo.SetFocus
   Exit Sub
End If
   adiciona_registrox
   suma_detalle
   Frame4.Visible = False
   Data2.Refresh
   found = ir_ultimo()
   Label40.Visible = False
   DBGrid2.SetFocus

End Sub

Private Sub Label3_Click()
Label40.Visible = True
Frame4.Visible = True
inicializa_xproducto
xcodigo.SetFocus

End Sub

Private Sub Label30_Click()
xteclado = "Producto"
Frame5.Visible = True

End Sub

Private Sub Label31_Click(Index As Integer)


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
               ticket_KeyPress 13
               Exit Sub
          Case 37 'delete
               ticket = ""
               Exit Sub
          Case 40 'close
               Frame5.Visible = False
               If Frame3.Visible = False Then Exit Sub
               ticket.SetFocus
               Exit Sub
   End Select
   ticket = ticket & Label31(Index)
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
If xteclado = "T1" Then
   Select Case Index
          Case 39 'emter
               Frame5.Visible = False
               If Frame6.Visible = False Then Exit Sub
               t1_KeyPress 13
               Exit Sub
          Case 37 'delete
               t1 = ""
               Exit Sub
          Case 40 'close
               Frame5.Visible = False
               If Frame6.Visible = False Then Exit Sub
               t1.SetFocus
               Exit Sub
   End Select
   t1 = t1 & Label31(Index)
End If
If xteclado = "T2" Then
   Select Case Index
          Case 39 'emter
               Frame5.Visible = False
               If Frame6.Visible = False Then Exit Sub
               t2_KeyPress 13
               Exit Sub
          Case 37 'delete
               t2 = ""
               Exit Sub
          Case 40 'close
               Frame5.Visible = False
               If Frame6.Visible = False Then Exit Sub
               t2.SetFocus
               Exit Sub
   End Select
   t2 = t2 & Label31(Index)
End If
If xteclado = "T3" Then
   Select Case Index
          Case 39 'emter
               Frame5.Visible = False
               If Frame6.Visible = False Then Exit Sub
               t3_KeyPress 13
               Exit Sub
          Case 37 'delete
               t3 = ""
               Exit Sub
          Case 40 'close
               Frame5.Visible = False
               If Frame6.Visible = False Then Exit Sub
               t3.SetFocus
               Exit Sub
   End Select
   t3 = t3 & Label31(Index)
End If
If xteclado = "T4" Then
   Select Case Index
          Case 39 'emter
               Frame5.Visible = False
               If Frame6.Visible = False Then Exit Sub
               t4_KeyPress 13
               Exit Sub
          Case 37 'delete
               t4 = ""
               Exit Sub
          Case 40 'close
               Frame5.Visible = False
               If Frame6.Visible = False Then Exit Sub
               t4.SetFocus
               Exit Sub
   End Select
   t4 = t4 & Label31(Index)
End If
If xteclado = "T5" Then
   Select Case Index
          Case 39 'emter
               Frame5.Visible = False
               If Frame6.Visible = False Then Exit Sub
               t5_KeyPress 13
               Exit Sub
          Case 37 'delete
               t5 = ""
               Exit Sub
          Case 40 'close
               Frame5.Visible = False
               If Frame6.Visible = False Then Exit Sub
               t5.SetFocus
               Exit Sub
   End Select
   t5 = t5 & Label31(Index)
End If
If xteclado = "T6" Then
   Select Case Index
          Case 39 'emter
               Frame5.Visible = False
               If Frame6.Visible = False Then Exit Sub
               t6_KeyPress 13
               Exit Sub
          Case 37 'delete
               t6 = ""
               Exit Sub
          Case 40 'close
               Frame5.Visible = False
               If Frame6.Visible = False Then Exit Sub
               t6.SetFocus
               Exit Sub
   End Select
   t6 = t6 & Label31(Index)
End If
If xteclado = "T7" Then
   Select Case Index
          Case 39 'emter
               Frame5.Visible = False
               If Frame6.Visible = False Then Exit Sub
               t7_KeyPress 13
               Exit Sub
          Case 37 'delete
               t7 = ""
               Exit Sub
          Case 40 'close
               Frame5.Visible = False
               If Frame6.Visible = False Then Exit Sub
               t7.SetFocus
               Exit Sub
   End Select
   t7 = t7 & Label31(Index)
End If






End Sub

Private Sub Label32_Click()

End Sub

Private Sub Label35_Click()
End Sub

Private Sub Label33_Click()
xteclado = "Busca"
Frame5.Visible = True

End Sub

Private Sub Label34_Change()
On Error GoTo cmd565_err
If Len("" & Data2.Recordset.Fields("linea")) > 0 Then
   Label34.Visible = True
   Else
   Label34.Visible = False
End If
Exit Sub
cmd565_err:
Exit Sub
End Sub

Private Sub Label34_Click()
On Error GoTo cmd123_err
If Len("" & Data2.Recordset.Fields("producto")) > 0 And Len("" & Data2.Recordset.Fields("linea")) > 0 Then
   ingreso_tallas "" & Data2.Recordset.Fields("linea")
   Exit Sub
End If
xcant = "" & Data2.Recordset.Fields("cantidad")
Frame6.Visible = True
xcant.SetFocus
Exit Sub
cmd123_err:
Data2.Refresh

End Sub

Private Sub Label39_Click()
Dim buf As String
Dim found As Integer
   adiciona_registro
   suma_detalle
   Frame1.Visible = False
   Data2.Refresh
   found = ir_ultimo()
   Label40.Visible = False
   DBGrid2.SetFocus
   DBGrid2.Col = 1

End Sub

Private Sub Label4_Click()
Label40.Visible = True
buffer = ""
Frame1.Visible = True
consulta_productos
DBGrid1.SetFocus
'Label33_Click

End Sub

Private Sub Label41_Click()
xteclado = "xcantidad"
Frame5.Visible = True

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
suma_detalle
Label49_Click
Exit Sub
cmd8911_err:
MsgBox "No se puede Grabar", 48, "Aviso"
Exit Sub

End Sub

Private Sub Label45_Click()
xteclado = "Cant"
Frame5.Visible = True


End Sub

Private Sub Label49_Click()
Dim found As Integer
Label40.Visible = False
Frame6.Visible = False
Data2.Refresh
DBGrid2.SetFocus
Frame5.Visible = False
found = ir_ultimo()

End Sub

Private Sub Label5_Click()
Dim found As Integer
borra_linea
found = suma_detalle()
Data2.Refresh

End Sub

Private Sub Label6_Click()
Dim found As Integer
If MsgBox("Desea Borrar", 1, "Aviso") <> 1 Then Exit Sub
borrar_todo
found = suma_detalle()

End Sub

Private Sub Label7_Click()
xteclado = "Clave"
Frame5.Visible = True

End Sub

Private Sub Label8_Click()
Dim found As Integer
If Len(ticket) = 0 Then
   ticket.SetFocus
   Exit Sub
End If
If vendedor = "*" Then
   vendedor.SetFocus
   Exit Sub
End If
'found = ingreso_existe()
'If found = 0 Then
'   MsgBox "No existe Ingreso", 48, "Aviso"
'   ticket.SetFocus
'   Exit Sub
'End If
If ingreso_existe() = 0 Then 'si ingreso el cliente
   MsgBox "No ha ingresado el cliente", 48, "Aviso"
   ticket.SetFocus
   Exit Sub
End If
cerrar_dataa
found = crear_temporal_pocket()
If found = 0 Then
   borrar_temporal
End If
found = seleccionar_pocket()
SQL_pedido   'visualizar sus pedidos por vendedor
Frame3.Visible = False
Frame5.Visible = False
found = suma_detalle()
found = ir_ultimo()
ulvendedor = vendedor


End Sub

Function crear_temporal_pocket()
On Error GoTo cmd2_err
borrar_archivo globaldir & "\_po" & ticket & ".dbf"
borrar_archivo globaldir & "\_po" & ticket & ".cdx"
FileCopy globaldir & "\tdetalle.dbf", globaldir & "\" & "_po" & ticket & ".dbf"
FileCopy globaldir & "\tdetalle.cdx", globaldir & "\" & "_po" & ticket & ".cdx"
crear_temporal_pocket = 1
Exit Function
cmd2_err:
Exit Function
End Function
Function seleccionar_pocket()
Dim i As Integer
Dim mytabley As Table
Dim mytablex As Table
Set mytabley = mydbxglo.OpenTable("_po" & ticket)
mytabley.Index = "tdetalle"
xnueo:
mytabley.Seek "=", "01", "PO", "001", ticket
If Not mytabley.NoMatch Then
   mytabley.Delete
   GoTo xnueo
End If
Set mytablex = mydbxglo.OpenTable("dproform")
mytablex.Index = "tdetalle"
mytablex.Seek "=", "01", "PO", "001", ticket
If Not mytablex.NoMatch Then
   Do
     If mytablex.EOF Then Exit Do
     If "" & mytablex.Fields("local") = "01" And "" & mytablex.Fields("tipo") = "PO" And "" & mytablex.Fields("serie") = "001" And "" & mytablex.Fields("numero") = ticket Then
        '-----------------------------------------
        mytabley.AddNew
        For i = 0 To mytablex.Fields.Count - 1
            mytabley.Fields(i) = mytablex.Fields(i)
        Next i
        mytabley.Fields("local") = "01"
        mytabley.Fields("tipo") = "PO"
        mytabley.Fields("serie") = "001"
        mytabley.Fields("numero") = ticket
        'mytabley.Fields("vendedor") = extra_loquesea(vendedor)
        mytabley.Update
        '-----------------------------------------
        Else: Exit Do
     End If
     mytablex.MoveNext
   Loop
End If
mytablex.Close
mytabley.Close
End Function
Function ingreso_existe()
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("ppocket")
mytablex.Index = "ppocket"
mytablex.Seek "=", ticket
If Not mytablex.NoMatch Then
   ingreso_existe = 1
End If
mytablex.Close
End Function
Function adiciona_cabeza()
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("cproform")
mytablex.Index = "tfactura"
mytablex.Seek "=", "01", "PO", "001", ticket
If Not mytablex.NoMatch Then
   mytablex.Edit
   mytablex.Fields("tipoclie") = "C"
   mytablex.Fields("local") = "01"
   mytablex.Fields("tipo") = "PO"
   mytablex.Fields("serie") = "001"
   mytablex.Fields("numero") = "" & ticket
   mytablex.Fields("codigo") = "" '& Codigo
   mytablex.Fields("nombre") = "" '& nombre
   mytablex.Fields("moneda") = "S" '& mytable11.Fields("moneda")
   mytablex.Fields("bodega") = "01" ' & mytable11.Fields("bodega")
   mytablex.Fields("estado") = "2"
   mytablex.Fields("usuario") = extra_loquesea("" & vendedor)
   mytablex.Fields("total") = Val("" & Total)
   mytablex.Fields("CAJA") = "T1" '& clave
   mytablex.Fields("tipoclie") = "C"
   mytablex.Fields("ACU") = "G"
   mytablex.Fields("SERVICIO") = "*"
   mytablex.Fields("vendedor") = "" 'extra_loquesea("" & vendedor)
   mytablex.Fields("observa") = "" '& observa
   mytablex.Fields("fechaCREA") = Format(Now, "dd/mm/yyyy")
   mytablex.Fields("fecha") = Format(Now, "dd/mm/yyyy")
   mytablex.Fields("fechaE") = Format(Now, "dd/mm/yyyy")
   mytablex.Fields("hora") = Format(Now, "hh:mm:ss")
   mytablex.Update
End If

If mytablex.NoMatch Then
   mytablex.AddNew
   mytablex.Fields("tipoclie") = "C"
   mytablex.Fields("local") = "01"
   mytablex.Fields("tipo") = "PO"
   mytablex.Fields("serie") = "001"
   mytablex.Fields("numero") = "" & ticket
   mytablex.Fields("codigo") = "" '& Codigo
   mytablex.Fields("nombre") = "" '& nombre
   mytablex.Fields("moneda") = "S" '& mytable11.Fields("moneda")
   mytablex.Fields("bodega") = "01" ' & mytable11.Fields("bodega")
   mytablex.Fields("estado") = "2"
   mytablex.Fields("usuario") = extra_loquesea("" & vendedor)
   mytablex.Fields("total") = Val("" & Total)
   mytablex.Fields("CAJA") = "T1" '& clave
   mytablex.Fields("tipoclie") = "C"
   mytablex.Fields("ACU") = "G"
   mytablex.Fields("SERVICIO") = "*"
   mytablex.Fields("vendedor") = "" 'extra_loquesea("" & vendedor)
   mytablex.Fields("observa") = "" '& observa
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
Dim rs
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("dproform")
mytablex.Index = "pocket"
apx1:
mytablex.Seek "=", "01", "PO", "001" ' , numero, extra_loquesea("" & vendedor)
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
   mytablex.Fields("local") = "01"
   mytablex.Fields("tipo") = "PO"
   mytablex.Fields("serie") = "001"
   mytablex.Fields("caja") = "T1"
   mytablex.Fields("numero") = ticket
   'mytablex.Fields("codigo") = "" & Codigo
   'mytablex.Fields("nombre") = "" & nombre
   'mytablex.Fields("vendedor") = extra_loquesea("" & vendedor)
   mytablex.Fields("fecha") = Format(Now, "dd/mm/yyyy")
   mytablex.Fields("hora") = Format(Now, "hh:mm:ss")
   mytablex.Fields("moneda") = "S" '& mytable11.Fields("moneda")
   mytablex.Fields("bodega") = "01" ' & mytable11.Fields("bodega")
   mytablex.Fields("estado") = "2"
   mytablex.Fields("usuario") = extra_loquesea("" & vendedor)
End Sub
Sub inicializa_xproducto()
xnombre = ""
xunidad = ""
xfactor = ""
xpventa = ""
xcantidad = "1"
xstock = ""
End Sub
Function busca_xproducto()
Dim mytablex As Table
Dim sw As Integer
xlinea = ""
sw = 0
Set mytablex = mydbxglo.OpenTable("producto")
mytablex.Index = "producto"
mytablex.Seek "=", xcodigo
If Not mytablex.NoMatch Then
   sw = 1
   xnombre = "" & mytablex.Fields("descripcio")
End If
mytablex.Close
If sw = 1 Then
Set mytablex = mydbxglo.OpenTable("precios")
mytablex.Index = "tprecios"
mytablex.Seek "=", xcodigo, "01"
If Not mytablex.NoMatch Then
   xlinea = "" & mytablex.Fields("linea")
   xunidad = "" & mytablex.Fields("unidad1")
   xfactor = "" & mytablex.Fields("factor1")
   xpventa = "" & mytablex.Fields("pventa1")
   busca_xproducto = 1
End If
mytablex.Close
End If
End Function


Private Sub nt1_Click()
xteclado = "T1"
Frame5.Visible = True

End Sub

Private Sub nt2_Click()
xteclado = "T2"
Frame5.Visible = True

End Sub

Private Sub nt3_Click()
xteclado = "T3"
Frame5.Visible = True

End Sub

Private Sub nt4_Click()
xteclado = "T4"
Frame5.Visible = True

End Sub

Private Sub nt5_Click()
xteclado = "T5"
Frame5.Visible = True

End Sub

Private Sub nt6_Click()
xteclado = "T6"
Frame5.Visible = True

End Sub

Private Sub nt7_Click()
xteclado = "T7"
Frame5.Visible = True

End Sub

Private Sub nt8_Click()
xteclado = "T8"
Frame5.Visible = True

End Sub

Private Sub t1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
t2.SetFocus
End Sub

Private Sub t2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
t3.SetFocus

End Sub

Private Sub t3_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
t4.SetFocus

End Sub

Private Sub t4_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
t5.SetFocus

End Sub

Private Sub t5_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
t6.SetFocus

End Sub

Private Sub t6_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
t7.SetFocus

End Sub

Private Sub t7_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
t8.SetFocus

End Sub

Private Sub ticket_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
vendedor.SetFocus
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
inicializa_xproducto
found = busca_xproducto()
If found = 0 Then
   MsgBox "NO existe Producto", 48, "Aviso"
   xcodigo = ""
   xcodigo.SetFocus
   Exit Sub
End If
xcantidad = "1"
xcantidad.SetFocus

End Sub
Sub borrar_temporal()
On Error GoTo cmd7812_err
mydbxglo.Execute "DELETE FROM " & "_po" & ticket
Exit Sub
cmd7812_err:
Exit Sub
End Sub
Sub cerrar_dataa()
On Error GoTo cmd43_err
Data2.Recordset.Close
Exit Sub
cmd43_err:
Exit Sub
End Sub
Sub sumar_totalg()
Dim mytablex As Table
Dim sdx As Double
sdx = 0

Set mytablex = mydbxglo.OpenTable("dproform")
mytablex.Index = "tdetalle"
mytablex.Seek "=", "01", "PO", "001", ticket
If Not mytablex.NoMatch Then
   Do
     If mytablex.EOF Then Exit Do
     If "" & mytablex.Fields("local") = "01" And "" & mytablex.Fields("tipo") = "PO" And "" & mytablex.Fields("serie") = "001" And "" & mytablex.Fields("numero") = ticket Then
        sdx = sdx + Val("" & mytablex.Fields("total"))
        Else: Exit Do
     End If
     mytablex.MoveNext
  Loop
End If
mytablex.Close
'
End Sub
Sub carga_familia()
Dim mytablex As Table
Dim i As Integer
For i = 0 To 14999
    mfamcod(i) = ""
Next i
i = -1
Set mytablex = mydbxglo.OpenTable("ppocket")
Do
If mytablex.EOF Then Exit Do
i = i + 1
mfamcod(i) = "" & mytablex.Fields("pedido")
mytablex.MoveNext
Loop
mfamtop = i
mytablex.Close
mfampag = 0
menu_familia "INI"

End Sub
Sub menu_familia(buf As String)
Dim i As Integer
Dim J As Integer
Select Case buf
       Case "INI"
            mfampag = 0
       Case "SIG"
            mfampag = mfampag + 29
            If mfampag > 300 Then
               mfampag = 0
            End If
       Case "ANT"
            mfampag = mfampag - 29
            If mfampag < 0 Then
               mfampag = 0
            End If
End Select
J = -1
For i = mfampag To 29 + mfampag
    J = J + 1
    Command1(J).Caption = mfamcod(i)
Next i

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
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("linea")
mytablex.Index = "linea"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   busca_linea = 1
   nlinea = "" & mytablex.Fields("descripcio")
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
'------------------------------------- ------------
mytablex.Close
End Function

Sub pone_tallas()
t1 = "" & Data2.Recordset.Fields("t1")
t2 = "" & Data2.Recordset.Fields("t2")
t3 = "" & Data2.Recordset.Fields("t3")
t4 = "" & Data2.Recordset.Fields("t4")
t5 = "" & Data2.Recordset.Fields("t5")
t6 = "" & Data2.Recordset.Fields("t6")
t7 = "" & Data2.Recordset.Fields("t7")
t8 = "" & Data2.Recordset.Fields("t8")
t9 = "" & Data2.Recordset.Fields("t9")
t10 = "" & Data2.Recordset.Fields("t10")
t11 = "" & Data2.Recordset.Fields("t11")
t12 = "" & Data2.Recordset.Fields("t12")
t13 = "" & Data2.Recordset.Fields("t13")
t14 = "" & Data2.Recordset.Fields("t14")
t15 = "" & Data2.Recordset.Fields("t15")
t16 = "" & Data2.Recordset.Fields("t16")
End Sub
'AQUI SI SE QUIERE GENERAR UNO NUEVO
Function busca_numero()
Dim mytablex As Table
Dim sdx As Double
ticket = ""
Set mytablex = mydbxglo.OpenTable("parame")
mytablex.Index = "codigo"
mytablex.Seek "=", "01"
If Not mytablex.NoMatch Then
   sdx = Val("" & mytablex.Fields("pocket")) + 1
   ticket = "" & sdx
End If
mytablex.Close
End Function
Function graba_numero()
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("parame")
mytablex.Index = "codigo"
mytablex.Seek "=", "01"
If Not mytablex.NoMatch Then
   mytablex.Edit
   mytablex.Fields("pocket") = Val(ticket)
   mytablex.Update
   graba_numero = 1
End If
mytablex.Close
End Function
Function valida_numero()
Dim mytablex As Table
Dim sdx As Double
Set mytablex = mydbxglo.OpenTable("Ppocket")
mytablex.Index = "ppocket"
contix:
mytablex.Seek "=", ticket
If Not mytablex.NoMatch Then
   sdx = Val(ticket) + 1
   ticket = "" & sdx
   GoTo contix
End If
mytablex.Close
End Function

Function prepara_achivo()
Dim found As Integer
Dim i As Integer
Dim sdx As Double
found = busca_numero()
found = valida_numero()
found = valida_cliente()
found = graba_numero()
Exit Function
End Function
Function valida_cliente()
Dim mytablex As Table
Dim found As Integer
On Error GoTo cmd1_err
Set mytablex = mydbxglo.OpenTable("ppocket")
mytablex.AddNew
mytablex.Fields("pedido") = "" & ticket
mytablex.Update
mytablex.Close
valida_cliente = 1
Exit Function
cmd1_err:
MsgBox "Error ,llame a servicio" + error$, 48, "Aviso"
Exit Function
End Function





