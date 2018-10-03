VERSION 5.00
Begin VB.Form trgb 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colores"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox tamano 
      Height          =   375
      Left            =   5760
      MaxLength       =   2
      TabIndex        =   7
      Text            =   "9"
      Top             =   1320
      Width           =   735
   End
   Begin VB.HScrollBar HS1 
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
   Begin VB.HScrollBar hs2 
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.HScrollBar hs3 
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tamaño"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label tipo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "          Guardar"
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label elcolor 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3615
   End
   Begin VB.Menu fdk99 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "trgb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub fdk99_Click()
    trgb.Hide
    Unload trgb

End Sub

Private Sub Form_Activate()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM paramecacolor where  caja='" & "" & mytable11.Fields("caja") & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        If tipo = "FAMILIA" Then

            HS1.Value = Val("" & mytablex.Fields("colorfamilia1"))
            hs2.Value = Val("" & mytablex.Fields("colorfamilia2"))
            hs3.Value = Val("" & mytablex.Fields("colorfamilia3"))
            tamano = Val("" & mytablex.Fields("sizefamilia"))

        End If

        If tipo = "PRODUCTO" Then
            HS1.Value = Val("" & mytablex.Fields("colorproducto1"))
            hs2.Value = Val("" & mytablex.Fields("colorproducto2"))
            hs3.Value = Val("" & mytablex.Fields("colorproducto3"))
            tamano = Val("" & mytablex.Fields("size"))

        End If

    End If

End Sub

Private Sub Form_Load()
    HS1.Min = 0
    HS1.max = 255
    HS1.LargeChange = 25
    HS1.SmallChange = 5

    hs2.Min = 0
    hs2.max = 255
    hs2.LargeChange = 25
    hs2.SmallChange = 5

    hs3.Min = 0
    hs3.max = 255
    hs3.LargeChange = 25
    hs3.SmallChange = 5

End Sub

Private Sub HS1_Change()
    elcolor.BackColor = RGB(HS1.Value, hs2.Value, hs3.Value)

End Sub

Private Sub hs2_Change()
    elcolor.BackColor = RGB(HS1.Value, hs2.Value, hs3.Value)

End Sub

Private Sub hs3_Change()
    elcolor.BackColor = RGB(HS1.Value, hs2.Value, hs3.Value)

End Sub

Private Sub Label1_Click()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM paramecacolor where  caja='" & "" & mytable11.Fields("caja") & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.AddNew
        mytablex.Fields("caja") = "" & mytable11.Fields("caja")

    End If

    If tipo = "FAMILIA" Then
        mytablex.Fields("colorfamilia1") = "" & HS1.Value
        mytablex.Fields("colorfamilia2") = "" & hs2.Value
        mytablex.Fields("colorfamilia3") = "" & hs3.Value
        mytablex.Fields("sizefamilia") = Val(tamano)
        mytablex.Update

    End If

    If tipo = "PRODUCTO" Then
        mytablex.Fields("colorproducto1") = "" & HS1.Value
        mytablex.Fields("colorproducto2") = "" & hs2.Value
        mytablex.Fields("colorproducto3") = "" & hs3.Value
        mytablex.Fields("size") = Val(tamano)
        mytablex.Update

    End If

    mytablex.Close
    fdk99_Click

End Sub
