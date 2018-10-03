VERSION 5.00
Begin VB.Form TXRXP 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Transmision Recepcion..."
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   12690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   12690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox campo 
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   4680
      Width           =   12495
   End
   Begin VB.ListBox List3 
      BackColor       =   &H00FFFFFF&
      Height          =   4155
      Left            =   8040
      TabIndex        =   7
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Procesar..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Todos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Borra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "<="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "=>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00FFFFFF&
      Height          =   4155
      Left            =   3720
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      Height          =   4155
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin VB.Menu li902 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "TXRXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub carga_basico()

    Dim sw As Integer

    List2.Clear
    List3.Clear

    Dim mytablex As ADODB.Recordset

    sw = 0
    mytablex.Open "select * from tlocal ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        sw = 1
        List2.AddItem "" & mytablex.Fields("nombre") & "|" & mytablex.Fields("codigo")
        mytablex.MoveNext
    Loop
    mytablex.Close

    If sw = 0 Then
        List2.ListIndex = 0

    End If

End Sub

