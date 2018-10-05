VERSION 5.00
Begin VB.Form tejecuta 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ejecucion de Comandos"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ejecuta"
      Height          =   615
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox comando 
      Height          =   1455
      Left            =   960
      MaxLength       =   200
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   240
      Width           =   6735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ALTER TABLE NombreTabla ADD NombreColumna NVARCHAR(20) NULL"
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   1800
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comando"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Menu flo2322 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tejecuta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    On Error GoTo cmd1_err

    cn.Execute (Comando)
    Exit Sub
cmd1_err:
    MsgBox "Error en formato " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub flo2322_Click()
    tejecuta.Hide
    Unload tejecuta

End Sub

