VERSION 5.00
Begin VB.Form trecibe 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recepcion de Tablas "
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Menu lo992 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "trecibe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lo992_Click()
trecibe.Hide
Unload trecibe
End Sub
