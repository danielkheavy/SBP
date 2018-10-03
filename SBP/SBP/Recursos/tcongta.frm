VERSION 5.00
Begin VB.Form tcongta 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuracion de la Conexion"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox clavesa 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   840
      Width           =   2655
   End
   Begin VB.ComboBox cmbservers 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar"
      Height          =   615
      Left            =   5760
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox servidor 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   30
      TabIndex        =   3
      Top             =   480
      Width           =   5175
   End
   Begin VB.TextBox clave 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clave Usuario SA"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Servidor"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clave Administrador"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Menu kfi333 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tcongta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbservers_Click()

    If Trim(cmbservers.Text) > 0 Then
        servidor = Trim(cmbservers.Text)

    End If

End Sub

Private Sub Command1_Click()

    Dim found As Integer

    If clave = "1967" Then
        If Len(servidor) = 0 Then Exit Sub
        If Len(servidor) > 30 Then Exit Sub
        found = grabar_servidor()

        If found = 0 Then
            MsgBox "No se pudo grabar ", 48, "Aviso"
            Exit Sub

        End If

        found = grabar_camino()

        If found = 0 Then
            MsgBox "No se pudo grabar ", 48, "Aviso"
            Exit Sub

        End If
   
        MsgBox "Proceso REalizado ", 48, "Aviso"
        Exit Sub

    End If

End Sub

Private Sub Form_Load()
    leer_servidor
    leer_camino

End Sub

Private Sub kfi333_Click()
    tcongta.Hide
    Unload tcongta

End Sub

Sub leer_servidor()

    Dim found As Integer

    Dim buf   As String

    On Error GoTo cmd169999_err

    'mirar_servidor
    servidor = ""
    buf = ""

    If Dir$(globalpath & "\server.txt") <> "" Then
        Close
        Open globalpath & "\server.txt" For Input As #1
        Input #1, buf
        Close #1
        servidor = buf

    End If

    Exit Sub
cmd169999_err:
    Close
    Exit Sub

End Sub

Sub leer_camino()

    Dim found As Integer

    Dim buf   As String

    On Error GoTo cmd16999912_err

    'mirar_servidor
    clavesa = ""
    buf = ""

    If Dir$(globalpath & "\camino.txt") <> "" Then
        Close
        Open globalpath & "\camino.txt" For Input As #1
        Input #1, buf
        Close #1
        clavesa = buf

    End If

    Exit Sub
cmd16999912_err:
    Close
    Exit Sub

End Sub

Function grabar_servidor()

    On Error GoTo cmd7823_err

    Dim buf As String

    Open globalpath & "\server.txt" For Output As #4
    buf = "" & servidor
    Print #4, buf;
    Close #4
    grabar_servidor = 1
    Exit Function
cmd7823_err:
    MsgBox "Aviso en graba servidor " + error$, 48, "Aviso"
    Exit Function

End Function

Function grabar_camino()

    On Error GoTo cmd78231_err

    Dim buf As String

    Open globalpath & "\camino.txt" For Output As #4
    buf = "" & clavesa
    Print #4, buf;
    Close #4
    grabar_camino = 1
    Exit Function
cmd78231_err:
    MsgBox "Aviso en graba clave " + error$, 48, "Aviso"
    Exit Function

End Function
