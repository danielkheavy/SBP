VERSION 5.00
Object = "{19BD1EA6-6E36-45BA-AEBD-BCF3093017CC}#11.0#0"; "GorditoButton.ocx"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form tkeyboar 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Teclado"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   13440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   13440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox precio 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Franklin Gothic Heavy"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   89
      Top             =   6720
      Width           =   2505
   End
   Begin VB.ComboBox familia 
      BackColor       =   &H80000004&
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
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   88
      Top             =   6840
      Width           =   5415
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   29
      Left            =   765
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   9930
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   42
      Left            =   13950
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   7185
      Width           =   1335
   End
   Begin VB.TextBox producto 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Franklin Gothic Heavy"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      MaxLength       =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11745
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   9240
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   30
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9870
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   31
      Left            =   3345
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9870
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   32
      Left            =   4665
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9870
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Cancel          =   -1  'True
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   33
      Left            =   6015
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9765
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   34
      Left            =   7365
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9660
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   35
      Left            =   2400
      MaskColor       =   &H00004080&
      TabIndex        =   6
      Top             =   8925
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   36
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8865
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   37
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8640
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   38
      Left            =   13785
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6345
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   1365
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   9375
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   2685
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   9375
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   4005
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   9375
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   5370
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   9360
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   5
      Left            =   6735
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   9345
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   6
      Left            =   8070
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   9285
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   7
      Left            =   9375
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   9270
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   8
      Left            =   10755
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   9225
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   9
      Left            =   12165
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   9045
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   11
      Left            =   750
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   8805
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   12
      Left            =   1755
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   10470
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   13
      Left            =   3015
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   10530
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   14
      Left            =   4620
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   10410
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   15
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   10320
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   16
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   10245
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   17
      Left            =   8745
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   10305
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   18
      Left            =   10260
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   10050
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   19
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   10260
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "Ñ"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   20
      Left            =   13530
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8700
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   21
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9810
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   10
      Left            =   945
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   11385
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   22
      Left            =   2265
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   11385
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   23
      Left            =   3585
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   11385
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   24
      Left            =   4725
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   11385
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   25
      Left            =   6165
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   11325
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   26
      Left            =   7635
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   11415
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   41
      Left            =   9300
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   11355
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   48
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   40
      Left            =   10785
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   11355
      Width           =   2655
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "LIMPIAR"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   39
      Left            =   1290
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   12825
      Width           =   2655
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "ESPACIADOR"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   28
      Left            =   3930
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   12825
      Width           =   6615
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00808080&
      Caption         =   "ENTER"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   27
      Left            =   9270
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   12645
      Width           =   3975
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   37
      Left            =   10605
      TabIndex        =   45
      Top             =   1050
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "9"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   29
      Left            =   60
      TabIndex        =   46
      Top             =   1050
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "1"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   34
      Left            =   6660
      TabIndex        =   47
      Top             =   1050
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "6"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   30
      Left            =   1380
      TabIndex        =   48
      Top             =   1050
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "2"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   33
      Left            =   5340
      TabIndex        =   49
      Top             =   1050
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "5"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":0070
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   38
      Left            =   11925
      TabIndex        =   50
      Top             =   1050
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "0"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":008C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   35
      Left            =   7965
      TabIndex        =   51
      Top             =   1050
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "7"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":00A8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   36
      Left            =   9285
      TabIndex        =   52
      Top             =   1050
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "8"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":00C4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   32
      Left            =   4020
      TabIndex        =   53
      Top             =   1050
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "4"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":00E0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   31
      Left            =   2700
      TabIndex        =   54
      Top             =   1050
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "3"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":00FC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   8
      Left            =   10605
      TabIndex        =   55
      Top             =   1935
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "O"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":0118
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   0
      Left            =   60
      TabIndex        =   56
      Top             =   1935
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "Q"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":0134
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   5
      Left            =   6660
      TabIndex        =   57
      Top             =   1935
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "Y"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":0150
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   1
      Left            =   1380
      TabIndex        =   58
      Top             =   1935
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "W"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":016C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   4
      Left            =   5340
      TabIndex        =   59
      Top             =   1935
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "T"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":0188
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   9
      Left            =   11925
      TabIndex        =   60
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "P"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":01A4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   6
      Left            =   7965
      TabIndex        =   61
      Top             =   1935
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "U"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":01C0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   7
      Left            =   9285
      TabIndex        =   62
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "I"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":01DC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   3
      Left            =   4020
      TabIndex        =   63
      Top             =   1935
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "R"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":01F8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   2
      Left            =   2700
      TabIndex        =   64
      Top             =   1935
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "E"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":0214
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   19
      Left            =   10605
      TabIndex        =   65
      Top             =   2820
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "L"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":0230
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   11
      Left            =   60
      TabIndex        =   66
      Top             =   2820
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "A"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":024C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   16
      Left            =   6660
      TabIndex        =   67
      Top             =   2820
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "H"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":0268
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   12
      Left            =   1380
      TabIndex        =   68
      Top             =   2835
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "S"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":0284
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   15
      Left            =   5340
      TabIndex        =   69
      Top             =   2820
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "G"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":02A0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   20
      Left            =   11925
      TabIndex        =   70
      Top             =   2820
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "Ñ"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":02BC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   17
      Left            =   7965
      TabIndex        =   71
      Top             =   2820
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "J"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":02D8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   18
      Left            =   9285
      TabIndex        =   72
      Top             =   2820
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "K"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":02F4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   14
      Left            =   4020
      TabIndex        =   73
      Top             =   2820
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "F"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":0310
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   13
      Left            =   2700
      TabIndex        =   74
      Top             =   2820
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "D"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":032C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   21
      Left            =   60
      TabIndex        =   75
      Top             =   3705
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "Z"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":0348
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   25
      Left            =   6660
      TabIndex        =   76
      Top             =   3705
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "N"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":0364
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   10
      Left            =   1380
      TabIndex        =   77
      Top             =   3705
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "X"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":0380
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   24
      Left            =   5340
      TabIndex        =   78
      Top             =   3705
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "B"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":039C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   40
      Left            =   10605
      TabIndex        =   79
      Top             =   3705
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "<"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":03B8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   26
      Left            =   7965
      TabIndex        =   80
      Top             =   3705
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "M"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":03D4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   41
      Left            =   9285
      TabIndex        =   81
      Top             =   3705
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":03F0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   23
      Left            =   4020
      TabIndex        =   82
      Top             =   3705
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "V"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":040C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   22
      Left            =   2700
      TabIndex        =   83
      Top             =   3705
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   "C"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":0428
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   990
      Index           =   39
      Left            =   60
      TabIndex        =   84
      Top             =   4605
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   1746
      BTYPE           =   4
      TX              =   "Limpiar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":0444
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   975
      Index           =   27
      Left            =   10620
      TabIndex        =   85
      Top             =   4605
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   1720
      BTYPE           =   4
      TX              =   "ENTER"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":0460
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   975
      Index           =   28
      Left            =   4035
      TabIndex        =   86
      Top             =   4605
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   1720
      BTYPE           =   4
      TX              =   "ESPACIADOR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":047C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn xopcioness 
      Height          =   900
      Index           =   42
      Left            =   11955
      TabIndex        =   87
      Top             =   75
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      BTYPE           =   4
      TX              =   ":"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tkeyboar.frx":0498
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin GorditoButton.Boton Comando 
      Height          =   750
      Index           =   0
      Left            =   9840
      TabIndex        =   90
      Top             =   5640
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1323
      PicturePosition =   0
      Caption         =   "0"
      UseGif          =   -1  'True
      BackColor       =   4210752
      ResalteColor    =   12632256
      PlayGif         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton Comando 
      Height          =   750
      Index           =   1
      Left            =   10635
      TabIndex        =   91
      Top             =   5640
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1323
      PicturePosition =   0
      Caption         =   "1"
      UseGif          =   -1  'True
      BackColor       =   4210752
      ResalteColor    =   12632256
      PlayGif         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton Comando 
      Height          =   750
      Index           =   2
      Left            =   11430
      TabIndex        =   92
      Top             =   5640
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1323
      PicturePosition =   0
      Caption         =   "2"
      UseGif          =   -1  'True
      BackColor       =   4210752
      ResalteColor    =   12632256
      PlayGif         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton Comando 
      Height          =   750
      Index           =   3
      Left            =   12225
      TabIndex        =   93
      Top             =   5640
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1323
      PicturePosition =   0
      Caption         =   "3"
      UseGif          =   -1  'True
      BackColor       =   4210752
      ResalteColor    =   12632256
      PlayGif         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton Comando 
      Height          =   750
      Index           =   4
      Left            =   9855
      TabIndex        =   94
      Top             =   6345
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1323
      PicturePosition =   0
      Caption         =   "4"
      UseGif          =   -1  'True
      BackColor       =   4210752
      ResalteColor    =   12632256
      PlayGif         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton Comando 
      Height          =   750
      Index           =   5
      Left            =   10635
      TabIndex        =   95
      Top             =   6345
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1323
      PicturePosition =   0
      Caption         =   "5"
      UseGif          =   -1  'True
      BackColor       =   4210752
      ResalteColor    =   12632256
      PlayGif         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton Comando 
      Height          =   750
      Index           =   6
      Left            =   11430
      TabIndex        =   96
      Top             =   6345
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1323
      PicturePosition =   0
      Caption         =   "6"
      UseGif          =   -1  'True
      BackColor       =   4210752
      ResalteColor    =   12632256
      PlayGif         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton Comando 
      Height          =   750
      Index           =   7
      Left            =   12225
      TabIndex        =   97
      Top             =   6345
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1323
      PicturePosition =   0
      Caption         =   "7"
      UseGif          =   -1  'True
      BackColor       =   4210752
      ResalteColor    =   12632256
      PlayGif         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton Comando 
      Height          =   750
      Index           =   8
      Left            =   9855
      TabIndex        =   98
      Top             =   7050
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1323
      PicturePosition =   0
      Caption         =   "8"
      UseGif          =   -1  'True
      BackColor       =   4210752
      ResalteColor    =   12632256
      PlayGif         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton Comando 
      Height          =   750
      Index           =   9
      Left            =   10635
      TabIndex        =   99
      Top             =   7050
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1323
      PicturePosition =   0
      Caption         =   "9"
      UseGif          =   -1  'True
      BackColor       =   4210752
      ResalteColor    =   12632256
      PlayGif         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton Comando 
      Height          =   750
      Index           =   10
      Left            =   12225
      TabIndex        =   100
      Top             =   7050
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1323
      PicturePosition =   0
      Caption         =   "Limpiar"
      UseGif          =   -1  'True
      BackColor       =   4210752
      ResalteColor    =   12632256
      PlayGif         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton Comando 
      Height          =   750
      Index           =   11
      Left            =   11430
      TabIndex        =   101
      Top             =   7050
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1323
      PicturePosition =   0
      Caption         =   "."
      UseGif          =   -1  'True
      BackColor       =   4210752
      ResalteColor    =   12632256
      PlayGif         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin VB.TextBox codigo 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Franklin Gothic Heavy"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11640
      TabIndex        =   104
      Top             =   6000
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FAMILIA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1700
      TabIndex        =   103
      Top             =   6240
      Width           =   2895
   End
   Begin VB.Label Label54 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PRECIO DE VENTA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7200
      TabIndex        =   102
      Top             =   6240
      Width           =   2535
   End
   Begin VB.Label flag 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   12240
      TabIndex        =   41
      Top             =   360
      Width           =   870
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFC0&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   3  'Dot
      BorderWidth     =   7
      DrawMode        =   5  'Not Copy Pen
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   4530
      Left            =   13500
      Top             =   1095
      Width           =   30
   End
   Begin VB.Menu flo444 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tkeyboar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Comando_Click(Index As Integer)

    If Index = 10 Then
        precio = ""
        Exit Sub

    End If

    precio = precio & Comando(Index).Caption

End Sub

Private Sub flo444_Click()
    tkeyboar.Hide
    Unload tkeyboar

End Sub

Private Sub Form_Load()

    '15/08/2018 Cambiar Descripcion de producto venta de ventas

    If busca_CambiaDescripcionVentas(2) = "S" Then 'SI PERMITE CAMBIAR DESCRIPCION
        tkeyboar.Height = 8655
        busca_correlativo (0)
        Carga_Familia

    End If
      
    '15/08/2018 Cambiar Descripcion de producto venta de ventas

End Sub

Private Sub producto_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
   
    Select Case FLAG

            '       Case "BUSCAPRODUCTO"
            '       tptovta.xbuscadescripcio = producto
        Case "CANTIDAD"

            If Val(producto) > 0 Then
                tptovta.DBGrid2.columns("cantidad") = Val(producto)

            End If

        Case "PRECIO"

            If IsNumeric(producto) Then
                tptovta.DBGrid2.columns("precio") = Val(producto)

            End If
      
            '15/08/2018 Cambiar Descripcion de producto venta de ventas
        Case "DESCRIPCION"
            'If IsNumeric(producto) Then
          
            If busca_CambiaDescripcionVentas(1) = "S" Then
                tptovta.DBGrid2.columns("descripcio") = producto
                Call RegistraProducto

            End If
          
            ' End If
            '15/08/2018 Cambiar Descripcion de producto venta de ventas
      
        Case "RUC"
            tptovta.xruc = producto

        Case "NOMBRE"
            tptovta.xnombre = producto

        Case "DIRECCION"
            tptovta.xdireccion = producto

        Case "GLOSA"
            tptovta.xdistrito = producto
        
        Case "CODIGO"
            tptovta.codigo = producto
            tptovta.busca_codigo_descuento (producto)

        Case "NOMBREP"
            tptovta.nombre = producto

        Case "YDIRECCION"
            tptovta.ydireccion = producto

        Case "CORREO"
            tptovta.correo = producto
        
        Case "DDIRECCION"
            tptovta.ddireccion = producto

        Case "DNOMBRE"
            tptovta.dnombre = producto

        Case "FECHANAC"
            tptovta.fechanac = producto

        Case "DREFERENCIA"
            tptovta.referencia = producto

        Case "TELEFONO"
            tptovta.telefono = producto

        Case "DCODIGO"
            tptovta.dcodigo = producto

        Case "TCAMPO1"
            tptovta.tcampo1 = producto
    
            If Len(tptovta.tcampo1) > 0 Then
                found = tptovta.busca_codigocl("" & tptovta.tcampo1, 0)

            End If

            tptovta.saldoabo = ""

            If "" & tptovta.dbgrid10.columns("tipo") = "C" And found = 1 Then
                found = tptovta.busca_credito_credito("" & tptovta.dbgrid10.columns("tipo"), "" & tptovta.tcampo1)  'actualiza su saldo actual

            End If
           
        Case "CONGELA"
            tptovta.xcongelax = producto
       
        Case "TCAMPO2"
            tptovta.tcampo2 = producto

        Case "TCAMPO3"
            tptovta.tcampo3 = producto

        Case "TCAMPO4"
            tptovta.tcampo4 = producto

        Case "HORA"

            If valida_hora("" & producto) = 1 Then
                tptovta.horaentrega = producto
                Else: MsgBox "Formato Hora No valido (HH:MM:SS)", 48, "Aviso"
                Exit Sub

            End If
       
    End Select

    flo444_Click
    Exit Sub

End Sub

Private Sub teclas_Click(Index As Integer)

    If Index = 39 Then
        producto = ""
        Exit Sub

    End If

    If Index = 27 Then
        producto_KeyPress 13
        Exit Sub

    End If

    If Index = 28 Then
        producto = producto & " "
        Exit Sub

    End If

    If Index = 40 Then
        If Len(producto) = 0 Then Exit Sub
        producto = Mid$(producto, 1, Len(producto) - 1)
        Exit Sub

    End If

    producto = producto & teclas(Index).Caption

End Sub

'15/08/2018 Cambiar Descripcion de producto venta de ventas
Private Sub xopcioness_Click(Index As Integer)

    If Index = 39 Then
        producto = ""
        Exit Sub

    End If

    If Index = 27 Then
        producto_KeyPress 13
        Exit Sub

    End If
     
    If Index = 28 Then
        producto = producto & " "
        Exit Sub

    End If

    If Index = 40 Then
        If Len(producto) = 0 Then Exit Sub
        producto = Mid$(producto, 1, Len(producto) - 1)
        Exit Sub

    End If
       
    producto = producto & teclas(Index).Caption

End Sub

'15/08/2018 Cambiar Descripcion de producto venta de ventas

'15/08/2018 Cambiar Descripcion de producto venta de ventas
Private Sub Carga_Familia()

    Dim cad      As String

    Dim mytablex As New ADODB.Recordset

    familia.Clear

    cad = "SELECT * FROM FAMILIA order by tipo desc "

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open cad, cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        familia.AddItem Trim("" & mytablex.Fields("descripcio")) & "|" & mytablex.Fields("familia")
        mytablex.MoveNext
    Loop
    familia.ListIndex = 0
    mytablex.Close

End Sub

'15/08/2018 Cambiar Descripcion de producto venta de ventas

'15/08/2018 Cambiar Descripcion de producto venta de ventas
Private Sub RegistraProducto()

    Dim found    As Integer

    Dim sdx      As Double

    Dim mytabley As New ADODB.Recordset

    Dim mytablex As New ADODB.Recordset

    Dim sw       As Integer

    sw = 0

    If Len(Trim(codigo)) = 0 Then
        Exit Sub

    End If
   
    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM producto where producto='" & Trim(codigo) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        mytablex.Close
        MsgBox "Ya existe codigo Usado ", 48, "Aviso"
        Exit Sub

    End If

    mytablex.AddNew
    grabando mytablex
    mytablex.Update
    graba_precios
    mytablex.Close
    'MsgBox "Producto Grabado ", 48, "Aviso"
    
    If tptovta.DBGrid2.Row = -1 Then
        MsgBox "No hay ningún registro para eliminar", vbInformation
        Exit Sub

    End If

    'If MsgBox("Se va a eliminar el registro : está seguro ", _
    '   vbExclamation + vbYesNo, "Eliminar") = vbYes Then
    productoCreado = codigo
    codigo = ""
    tptovta.Data2.Recordset.Delete

    If tptovta.Data2.Recordset.EOF = True And tptovta.Data2.Recordset.BOF = True Then
        Exit Sub

    End If

    found = tptovta.sumar_detalle()
    ' End If
    Exit Sub

End Sub

'15/08/2018 Cambiar Descripcion de producto venta de ventas

'15/08/2018 Cambiar Descripcion de producto venta de ventas
Sub grabando(mytablex As ADODB.Recordset)

    Dim found As Integer

    fotonombre = ""
    mytablex.Fields("fechavence") = Null

    mytablex.Fields("diasalerta") = ""
    mytablex.Fields("seinventaria") = ""
    mytablex.Fields("tecla") = ""
    unidadp = "UND"
    factorp = 1

    mytablex.Fields("dia") = ""
    mytablex.Fields("fueldonde") = ""
    mytablex.Fields("comisioncredito") = 0
    mytablex.Fields("costoanterior1") = 0
    mytablex.Fields("costoanterior2") = 0

    mytablex.Fields("puertoimpresion1") = ""
    mytablex.Fields("puertoimpresion2") = ""
    mytablex.Fields("puertoimpresion3") = ""

    mytablex.Fields("puertoimpresion") = "COCINA"
    mytablex.Fields("grupoimpresion") = "C"
    mytablex.Fields("cola") = "S"
    mytablex.Fields("recetaprn") = ""
    mytablex.Fields("empaque_visible") = ""
    mytablex.Fields("platos") = 0
    mytablex.Fields("fuel") = ""
    mytablex.Fields("touch") = 0
    mytablex.Fields("dsctoref") = 0
    mytablex.Fields("unidadp") = "UND"
    mytablex.Fields("factorp") = "1"
    mytablex.Fields("margen") = ""
    mytablex.Fields("OK") = ""
    mytablex.Fields("percepcion") = ""
    mytablex.Fields("producto") = Trim(codigo)
    mytablex.Fields("detraccion") = 0
    mytablex.Fields("ivap") = 0
    mytablex.Fields("flete") = 0
    mytablex.Fields("barras") = ""
    mytablex.Fields("descripcio") = UCase$(Trim(producto))
    mytablex.Fields("descorto") = Mid$(producto, 1, 22)
    mytablex.Fields("presenta") = ""
    mytablex.Fields("familia") = extra_loquesea1(Trim(familia))
    mytablex.Fields("subfamilia") = ""
    mytablex.Fields("seccion") = ""
    mytablex.Fields("marca") = ""
    mytablex.Fields("remate") = "N"

    mytablex.Fields("tipocreacion") = "NUEVO"
    mytablex.Fields("obligacomentario") = "N"
    mytablex.Fields("CostoReceta") = "S"
    mytablex.Fields("categoria") = ""
    mytablex.Fields("linea") = ""
    mytablex.Fields("color") = ""

    mytablex.Fields("talla") = ""
    mytablex.Fields("proyecto") = ""
    mytablex.Fields("sexo") = ""
    mytablex.Fields("procedencia") = ""

    mytablex.Fields("fabrica") = ""
    mytablex.Fields("serviciomesa") = 0

    mytablex.Fields("serie") = "N"
    mytablex.Fields("peso") = "N"
    mytablex.Fields("servicio") = "N"
    mytablex.Fields("vtaund") = "S"
    mytablex.Fields("oferta") = "N"
    mytablex.Fields("vecaja") = "S"
    mytablex.Fields("estado") = "S"

    mytablex.Fields("igv") = busca_ValorIgv()

    mytablex.Fields("isc") = 0
    mytablex.Fields("pesokgr") = 0
    mytablex.Fields("comision") = 0
    mytablex.Fields("monedac") = "S"
    mytablex.Fields("unidad") = "UND"
    mytablex.Fields("factor") = 1
    mytablex.Fields("costou") = 0
    mytablex.Fields("costop") = 0
    mytablex.Fields("costoini") = 0
    mytablex.Fields("minimo") = 0
    mytablex.Fields("maximo") = 0

    mytablex.Fields("monedav") = s
    mytablex.Fields("cospaqu") = 0
    mytablex.Fields("cospaqp") = 0
    mytablex.Fields("cospaqi") = 0

    Exit Sub
cmd7832_err:
    MsgBox "Aviso en grabando " + error$, 48, "Aviso"
    Exit Sub

End Sub

'15/08/2018 Cambiar Descripcion de producto venta de ventas
    
Sub graba_precios()

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    mytablex.Open "SELECT * FROM precios where  producto='" & Trim(codigo) & "' ", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.AddNew
        mytablex.Fields("producto") = Trim(codigo)
        mytablex.Fields("local") = "01"
        mytablex.Fields("UNIDAD1") = "UND"
        mytablex.Fields("FACTOR1") = "1"
        mytablex.Fields("PVENTA1") = Val(precio)
        mytablex.Fields("DSCTO") = 0
        ' graba_xprecio mytablex
        mytablex.Update

    End If

    mytablex.Close
    Exit Sub

End Sub
    
'15/08/2018 Cambiar Descripcion de producto venta de ventas
Sub busca_correlativo(sw As Integer)

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    If sw = 0 Then
        sdx = Val("" & mytablex.Fields("producto")) + 1
        codigo = "" & sdx

    End If

    mytablex.Close
sigueb:
    mytablex.Open "select * from producto where producto='" & Trim(codigo) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        mytablex.Close
        sdx = sdx + 1
        codigo = "" & sdx
        GoTo sigueb
        Exit Sub

    End If

End Sub

'15/08/2018 Cambiar Descripcion de producto venta de ventas

'15/08/2018 Cambiar Descripcion de producto venta de ventas
Function busca_ValorIgv()

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select igv from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_ValorIgv = ("" & mytablex.Fields("igv"))

    End If

    mytablex.Close
    Exit Function

End Function

'15/08/2018 Cambiar Descripcion de producto venta de ventas
