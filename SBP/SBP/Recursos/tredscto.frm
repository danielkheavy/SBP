VERSION 5.00
Begin VB.Form Tredscto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Descuento/recargas"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   8430
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox valor 
      Height          =   495
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label total 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Digite un Valor, presione entern"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Menu flo89923 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "Tredscto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dj782321_Click()
    flo89923_Click

End Sub

Private Sub flo89923_Click()

    Tredscto.Hide
    Unload Tredscto

End Sub

Private Sub Form_Load()
    List1.Clear
    List1.AddItem "DESCUENTO EN % -TOTAL "
    List1.AddItem "DESCUENTO EN MONTO - TOTAL"
    List1.AddItem "BORRAR TODOS LOS DESCUENTOS TOTALES"
    List1.AddItem "RECARGO EN % -TOTAL "
    List1.ListIndex = 0

End Sub

Private Sub List1_DblClick()
    List1_KeyPress 13

End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        flo89923_Click
        Exit Sub

    End If

    valor.SetFocus

End Sub

Private Sub valor_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        flo89923_Click
        Exit Sub

    End If

    If List1.ListIndex = 1 Then
        If Val(valor) > Val(total) Then
            MsgBox "Rango Descuentos Erroneo", 24, "Aviso"
            valor.SetFocus
            Exit Sub

        End If

    End If

    If Not IsNumeric(valor) Then
        valor = ""
        valor.SetFocus
        Exit Sub

    End If

    tipodescuento = "" & List1.ListIndex
       
    valordescuento = valor
    flo89923_Click

End Sub

