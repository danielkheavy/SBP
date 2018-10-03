VERSION 5.00
Begin VB.Form tcajade 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Caja o Terminal Defecto"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   13065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox puerto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar"
      Height          =   735
      Left            =   2640
      TabIndex        =   0
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nro de Caja o Terminal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.Menu floo312 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tcajade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type bdcaja

    ppuerto As String * 10

End Type

Private Sub Command1_Click()

    Dim found As Integer

    If Len(Trim(Puerto)) = 0 Then
        Puerto.SetFocus

    End If

    found = grabar_visor()

End Sub

Private Sub floo312_Click()
    tcajade.Hide
    Unload tcajade

End Sub

Private Sub Form_Load()

    Dim found As Integer

    found = leer_visor()

End Sub

Function grabar_visor()

    Dim bdcaja1 As bdcaja

    On Error GoTo cmd7823_err

    Dim buf As String

    bdcaja1.ppuerto = Trim(Puerto)
    Open globalpath & "\caja.txt" For Random As #4 Len = Len(bdcaja1)
    Put #4, 1, bdcaja1
    Close #4
    MsgBox "Proceso Grabado ", 48, "Aviso"
    grabar_visor = 1
    Exit Function
cmd7823_err:
    MsgBox "Aviso en graba Caja " + error$, 48, "Aviso"
    Exit Function

End Function

Function leer_visor()

    Dim bdcaja1 As bdcaja

    On Error GoTo cmd7824_err

    Dim buf As String

    Open globalpath & "\caja.txt" For Random As #4 Len = Len(bdcaja1)
    Get #4, 1, bdcaja1
    Puerto = Trim(bdcaja1.ppuerto)
    leer_visor = 1
    Close #4
    Exit Function
cmd7824_err:
    MsgBox "Aviso en Leer Puerto " + error$, 48, "Aviso"
    Exit Function

End Function

