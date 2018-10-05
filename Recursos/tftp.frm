VERSION 5.00
Begin VB.Form tftp 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferencia Data "
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Recibir"
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Enviar"
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Menu lfo343 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tftp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
