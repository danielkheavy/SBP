VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End Function
Public Function conectar_remoto()
 On Error GoTo cmd5454_err
 Dim cnremote As New ADODB.Connection
 cnremote.CursorLocation = adUseClient
 cnremote.Open "Driver={SQL Server};Server=192.168.1.53;Database=calipso;Uid=sa"
 conectar_remoto = 1
 'MsgBox "Conexion buena", 48, "Aviso"
 Exit Function
cmd5454_err:
 Exit Function
 End Function

