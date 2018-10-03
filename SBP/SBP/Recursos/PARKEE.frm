VERSION 5.00
Begin VB.Form PARKEE 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrada de vehiculo"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   10035
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "GRABA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5880
      Top             =   6600
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   4560
      ScaleHeight     =   1035
      ScaleWidth      =   3795
      TabIndex        =   40
      Top             =   1080
      Width           =   3855
      Begin VB.Label fecha 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   41
         Top             =   120
         Width           =   2535
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   0
         Picture         =   "PARKEE.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   915
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   4560
      ScaleHeight     =   1035
      ScaleWidth      =   3795
      TabIndex        =   38
      Top             =   2280
      Width           =   3855
      Begin VB.Label hora 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   39
         Top             =   120
         Width           =   2535
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   1335
         Left            =   0
         Picture         =   "PARKEE.frx":0A4B
         Stretch         =   -1  'True
         Top             =   0
         Width           =   915
      End
   End
   Begin VB.TextBox PLACA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      MaxLength       =   20
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Z"
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
      Index           =   35
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Y"
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
      Index           =   34
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "X"
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
      Index           =   33
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "W"
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
      Index           =   32
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "U"
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
      Index           =   31
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "T"
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
      Index           =   30
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "S"
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
      Index           =   29
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "R"
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
      Index           =   28
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Q"
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
      Index           =   27
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "P"
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
      Index           =   26
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "O"
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
      Index           =   25
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "N"
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
      Index           =   24
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "M"
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
      Index           =   23
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "L"
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
      Index           =   22
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "K"
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
      Index           =   21
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "J"
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
      Index           =   20
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "I"
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
      Index           =   19
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "H"
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
      Index           =   18
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "G"
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
      Index           =   17
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "F"
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
      Index           =   16
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "E"
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
      Index           =   15
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "D"
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
      Index           =   14
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "C"
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
      Index           =   13
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "B"
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
      Index           =   12
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "A"
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
      Index           =   11
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   10
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "0"
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
      Index           =   9
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "9"
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
      Index           =   8
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "8"
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
      Index           =   7
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "7"
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
      Index           =   6
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "6"
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
      Index           =   5
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "5"
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
      Index           =   4
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "4"
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
      Index           =   3
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "3"
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
      Index           =   2
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "2"
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
      Index           =   1
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "1"
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
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   3255
      Left            =   4560
      Picture         =   "PARKEE.frx":1AE7
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   5415
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PLACA NRO."
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
      Left            =   4560
      TabIndex        =   37
      Top             =   240
      Width           =   1935
   End
   Begin VB.Menu ffsa 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "PARKEE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)
    placa = placa + Command1(Index).Caption

End Sub

Private Sub Command2_Click()

    Dim found As Integer

    If Len(placa) = 0 Then
        MsgBox "Ingrese numero de Placa", 48, "Aviso"
        placa.SetFocus
        Exit Sub

    End If

    found = graba_parqueo()

    If found = 1 Then
        proceso_formatospa
        ffsa_Click
        Exit Sub

    End If

    placa.SetFocus

End Sub

Private Sub Command3_Click()
    ffsa_Click

End Sub

Private Sub ffsa_Click()
    PARKEE.Hide
    Unload PARKEE

End Sub

Private Sub Label1_Click()
    placa = ""

End Sub

Private Sub Timer1_Timer()
    fecha = Format(Now, "dd/mm/yyyy")
    hora = Format(Now, "hh:mm:ss")

End Sub

Function graba_parqueo()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from movplaca where placa='" & placa & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        MsgBox "Ya Existe numero de Placa Ingresado ", 48, "Aviso"
        mytablex.Close
        Exit Function

    End If

    If MsgBox("Desea Grabar", 1, "Aviso") = 1 Then
        mytablex.AddNew
        mytablex.Fields("placa") = "" & placa
        mytablex.Fields("fecha") = Format(Now, "dd/mm/yyyy")
        mytablex.Fields("hora") = Format(Now, "hh:mm:ss")
        mytablex.Fields("vendedor") = gusuario
        mytablex.Update
        MsgBox "Transaccion Grabado", 48, "Aviso"
        graba_parqueo = 1

    End If

    mytablex.Close

End Function

Sub proceso_formatospa()

    Dim mytablex        As New ADODB.Recordset

    Dim archivo_formato As String

    Dim bxlocal         As String

    Dim bxtipo          As String

    Dim bxserie         As String

    Dim bxnumero        As String

    Dim found           As Integer

    bxlocal = ""
    bxtipo = ""
    bxserie = ""
    bxnumero = ""
    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    found = borra_nombre("" & FileName)
       
    archivo_formato = busca_archivo_formato()

    If Len(archivo_formato) = 0 Then
        MsgBox "No existe archivo formato ", 48, "Aviso"
        Exit Sub

    End If

    mytablex.Open "SELECT * FROM  movplaca   where  placa='" & placa & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then 'si existe
        mytablex.Close
        Exit Sub

    End If

    'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    'found = proceso_formatos(archivo_formato, mytablex, "{", "}", "MOVPLACA", "MOVPLACA", bxlocal, bxtipo, bxserie, bxnumero, "0", 0)
    found = proceso_formatos(archivo_formato, mytablex, "{", "}", "MOVPLACA", "MOVPLACA", bxlocal, bxtipo, bxserie, bxnumero, "0", 0)
    'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    mytablex.Close
    imprime_precuenta
    Exit Sub
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Function busca_archivo_formato() As String
    busca_archivo_formato = "parkeo"
 
End Function

Sub imprime_precuenta()

    Dim found As Integer

    Dim sFile As String

    'impresora por default atachado
    On Error GoTo cmd90000_err

    cerrar_archivo

    If Len(Trim("" & mytable11.Fields("ecpuerto"))) = 0 Then
        MsgBox "Puerto de Precuenta no configurado", 48, "Aviso"
        Exit Sub

    End If

    If Trim("" & mytable11.Fields("eccola")) = "S" Then
        sFile = globaldir & "\temporal\" & gusuario & ".txt"
        found = Imprime_archivojj(sFile, 0, "" & mytable11.Fields("tamanoletra"), "" & mytable11.Fields("nombrefont"), "" & mytable11.Fields("BOLD"), "" & mytable11.Fields("letrainterna"))
        Exit Sub

    End If

    If Trim("" & mytable11.Fields("eccola")) <> "S" Then
        found = star_sp342(Trim("" & mytable11.Fields("ecpuerto")), 0)
        found = corte_papel(Trim("" & mytable11.Fields("ecpuerto")), Val("" & mytable11.Fields("catipo")))
        Exit Sub

    End If

    Exit Sub
cmd90000_err:
    MsgBox "Error en imprime precuenta" + error$, 48, "Aviso"
    Exit Sub
    Exit Sub

End Sub

