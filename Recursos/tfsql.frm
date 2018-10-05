VERSION 5.00
Begin VB.Form tfsql 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prueba del Sql"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox cadena 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   7215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ejecutar"
      Height          =   615
      Left            =   5595
      TabIndex        =   0
      Top             =   4320
      Width           =   1695
   End
End
Attribute VB_Name = "tfsql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    cn.Execute (cadena)

End Sub

