VERSION 5.00
Begin VB.Form thiscli 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema de Historias Clinicas"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   11460
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFF00&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   11400
      TabIndex        =   0
      Top             =   0
      Width           =   11460
      Begin VB.CommandButton image2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Historia"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1920
         Picture         =   "thiscli.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton image3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&FichaPaciente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         Picture         =   "thiscli.frx":2682
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Menu dm882 
      Caption         =   "&Menu"
      Begin VB.Menu dk21 
         Caption         =   "&1.Ficha de Clientes"
      End
   End
   Begin VB.Menu dflo22 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "thiscli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dflo22_Click()
thiscli.Hide
Unload thiscli

End Sub
