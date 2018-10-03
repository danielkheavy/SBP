VERSION 5.00
Begin VB.Form tvisorc 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parametros Generales Visor"
   ClientHeight    =   6630
   ClientLeft      =   150
   ClientTop       =   750
   ClientWidth     =   13965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   13965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox velocidad 
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
      Left            =   3480
      MaxLength       =   30
      TabIndex        =   7
      Top             =   840
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar"
      Height          =   735
      Left            =   3480
      TabIndex        =   5
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox mensaje2 
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
      Left            =   3480
      MaxLength       =   30
      TabIndex        =   4
      Top             =   2280
      Width           =   7815
   End
   Begin VB.TextBox mensaje1 
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
      Left            =   3480
      MaxLength       =   30
      TabIndex        =   3
      Top             =   1560
      Width           =   7815
   End
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
      Left            =   3480
      MaxLength       =   30
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Velocidad (9600,n,8,1)"
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
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mensaje"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Puerto (1,2,3,4-)"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Menu lfo4344 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tvisorc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Dim found As Integer

    If Len(Trim(Puerto)) = 0 Then
        Puerto.SetFocus

    End If

    If Len(Trim(velocidad)) = 0 Then
        velocidad.SetFocus

    End If

    If Len(Trim(mensaje1)) = 0 Then
        mensaje1.SetFocus

    End If

    found = grabar_visor()

End Sub

Private Sub Form_Load()

    Dim found As Integer

    found = leer_visor()

End Sub

Private Sub Label3_Click()
    velocidad = "9600,n,8,1"

End Sub

Private Sub lfo4344_Click()
    tvisorc.Hide
    Unload tvisorc

End Sub

Function grabar_visor()

    Dim bdvisor1 As bdvisor

    On Error GoTo cmd7823_err

    Dim buf As String

    bdvisor1.ppuerto = Trim(Puerto)
    bdvisor1.vvelocidad = Trim(velocidad)
    bdvisor1.mmensaje1 = Trim(mensaje1)
    bdvisor1.mmensaje2 = Trim(mensaje2)
    Open globalpath & "\visor.txt" For Random As #4 Len = Len(bdvisor1)
    Put #4, 1, bdvisor1
    Close #4
    MsgBox "Proceso Grabado ", 48, "Aviso"
    grabar_visor = 1
    Exit Function
cmd7823_err:
    MsgBox "Aviso en graba visor " + error$, 48, "Aviso"
    Exit Function

End Function

Function leer_visor()

    Dim bdvisor1 As bdvisor

    On Error GoTo cmd7824_err

    Dim buf As String

    Open globalpath & "\visor.txt" For Random As #4 Len = Len(bdvisor1)
    Get #4, 1, bdvisor1
    Puerto = Trim(bdvisor1.ppuerto)
    velocidad = Trim(bdvisor1.vvelocidad)
    mensaje1 = Trim(bdvisor1.mmensaje1)
    mensaje2 = Trim(bdvisor1.mmensaje2)
    leer_visor = 1
    Close #4
    Exit Function
cmd7824_err:
    MsgBox "Aviso en Leer visor " + error$, 48, "Aviso"
    Exit Function

End Function
