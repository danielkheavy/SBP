VERSION 5.00
Begin VB.Form tdremoto 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Call Center Delivery"
   ClientHeight    =   8415
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   13515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   13515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton nlocales 
      BackColor       =   &H00C0C0C0&
      Height          =   855
      Index           =   9
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7440
      Width           =   4335
   End
   Begin VB.CommandButton nlocales 
      BackColor       =   &H00C0C0C0&
      Height          =   855
      Index           =   8
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6600
      Width           =   4335
   End
   Begin VB.CommandButton nlocales 
      BackColor       =   &H00C0C0C0&
      Height          =   855
      Index           =   7
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5760
      Width           =   4335
   End
   Begin VB.CommandButton nlocales 
      BackColor       =   &H00C0C0C0&
      Height          =   855
      Index           =   6
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4920
      Width           =   4335
   End
   Begin VB.CommandButton nlocales 
      BackColor       =   &H00C0C0C0&
      Height          =   855
      Index           =   5
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4080
      Width           =   4335
   End
   Begin VB.CommandButton nlocales 
      BackColor       =   &H00C0C0C0&
      Height          =   855
      Index           =   4
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7440
      Width           =   4335
   End
   Begin VB.CommandButton nlocales 
      BackColor       =   &H00C0C0C0&
      Height          =   855
      Index           =   3
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6600
      Width           =   4335
   End
   Begin VB.CommandButton nlocales 
      BackColor       =   &H00C0C0C0&
      Height          =   855
      Index           =   2
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5760
      Width           =   4335
   End
   Begin VB.CommandButton nlocales 
      BackColor       =   &H00C0C0C0&
      Height          =   855
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4920
      Width           =   4335
   End
   Begin VB.CommandButton nlocales 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4080
      Width           =   4335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Close"
      Height          =   615
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clear"
      Height          =   615
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1320
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Grabar"
      Height          =   615
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox fechanac 
      Height          =   495
      Left            =   2400
      MaxLength       =   10
      TabIndex        =   10
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox referencia 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   2400
      MaxLength       =   60
      TabIndex        =   4
      Top             =   2040
      Width           =   8175
   End
   Begin VB.TextBox distrito 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   2400
      MaxLength       =   15
      TabIndex        =   3
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox direccion 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   2400
      MaxLength       =   60
      TabIndex        =   2
      Top             =   1080
      Width           =   8175
   End
   Begin VB.TextBox nombre 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   2400
      MaxLength       =   60
      TabIndex        =   1
      Top             =   600
      Width           =   8175
   End
   Begin VB.TextBox telefono 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   2400
      MaxLength       =   11
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label local1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2400
      TabIndex        =   15
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cod Local"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label nlocal1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2400
      TabIndex        =   13
      Top             =   3000
      Width           =   8175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local  a Enviar"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Nacimiento"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Referencia"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Distrito"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Direccion"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Apellidos Nombres"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Telefono"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2295
   End
   Begin VB.Menu lo990 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tdremoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command3_Click()
    lo990_Click

End Sub

Private Sub direccion_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    distrito.SetFocus

End Sub

Private Sub distrito_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    distrito.SetFocus

End Sub

Private Sub fechanac_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub Form_Load()
    carga_locales

End Sub

Sub carga_locales()

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    I = 0
    mytablex.Open "select * from  tlocal", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        nlocales(I).Caption = "" & mytablex.Fields("nombre")
        I = I + 1
        mytablex.MoveNext
    Loop
    mytablex.Close

End Sub

Private Sub lo990_Click()
    tdremoto.Hide
    Unload tdremoto

End Sub

Private Sub nombre_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    direccion.SetFocus

End Sub

Private Sub referencia_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    fechanac.SetFocus

End Sub

Private Sub telefono_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    nombre.SetFocus

End Sub
