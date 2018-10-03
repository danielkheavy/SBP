VERSION 5.00
Begin VB.Form tteclado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Teclado"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   13275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox producto 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      MaxLength       =   15
      TabIndex        =   41
      Top             =   0
      Width           =   9255
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00C0C0C0&
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4200
      Width           =   3975
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00C0C0C0&
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4200
      Width           =   6615
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
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
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H000000FF&
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
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   2655
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00C0C0C0&
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   3975
   End
   Begin VB.Label xcontrol 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dato Digitado"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "tteclado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub teclas_Click(Index As Integer)

    If Index = 39 Then
          
        Exit Sub

    End If

    If Index = 27 Then
          
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
