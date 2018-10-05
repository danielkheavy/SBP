VERSION 5.00
Begin VB.Form Sqlserv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prueba sql Server"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox nombre 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox codigo 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   45
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Sqlserv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Connect As New ADODB.Connection
Function coneccion()
Dim chaine As String
chaine = "Provider=SQLOLEDB.1" & _
        ";Integrated Security=SSPI" & _
        ";Persist Security Info=False" & _
        ";Initial Catalog=NOM_BDD" & _
        ";Data Source=NOM_SERVEUR_BD"
On Error GoTo ERREUR
'link connection and string
Connect.ConnectionString = chaine
Connect.Open
coneccion = 1
Exit Function
ERREUR:
End Function



Private Sub Form_Load()
Dim found As Integer
found = coneccion()
If found = 0 Then
   MsgBox "No puedo Conectarse a la base de datos", 48, "Aviso"
   End
End If
MsgBox "Bueno"
End Sub

Private Sub Form_Unload(Cancel As Integer)
' Request of confirmation for
'unload application and disconnect data base
If MsgBox("Desea terminar la Aplicacion ", vbOKCancel + vbExclamation, "Quitter" & " L'application") = vbOK Then
   Connect.Close
End
      
Else
  Cancel = -1 'cancel=-1, Form_Unload is cancel
End If

End Sub
