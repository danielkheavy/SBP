VERSION 5.00
Begin VB.Form tdiasema 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Habilita Deshabilita Semanas"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Deshabilita Semana Programada"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   4815
   End
   Begin VB.Menu fdlo33 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tdiasema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    If MsgBox("esta Seguro", 1, "Aviso") <> 1 Then Exit Sub
    cn.Execute ("update producto set dia=''")
    MsgBox "Proceso Realizado ", 48, "Aviso"
    fdlo33_Click

End Sub

Private Sub fdlo33_Click()
    tdiasema.Hide
    Unload tdiasema

End Sub
