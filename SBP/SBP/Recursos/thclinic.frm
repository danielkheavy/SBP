VERSION 5.00
Begin VB.Form thclinic 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Control de Historias Clinicas"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   11085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton image5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Asistencia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2160
      Picture         =   "thclinic.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton image3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Ficha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      Picture         =   "thclinic.frx":21BA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton image2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Consulta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2160
      Picture         =   "thclinic.frx":3F78
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Tratamiento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      Picture         =   "thclinic.frx":65FA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton image14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Diagnostico"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3960
      MaskColor       =   &H8000000E&
      Picture         =   "thclinic.frx":9144
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin VB.Menu dk88225 
      Caption         =   "&Menu"
      Begin VB.Menu cli44 
         Caption         =   "&1.Clinicas"
      End
      Begin VB.Menu dksef 
         Caption         =   "&2.Cias de Seguro"
      End
      Begin VB.Menu refr664 
         Caption         =   "&3.Referencias"
      End
      Begin VB.Menu medi8844 
         Caption         =   "&4.Medicos"
      End
      Begin VB.Menu d7733 
         Caption         =   "&5.TipoAfiliacion"
      End
      Begin VB.Menu dk8822 
         Caption         =   "&6.Tipo Autorizacion"
      End
      Begin VB.Menu djw221 
         Caption         =   "&7.Tipo de Consulta"
      End
      Begin VB.Menu engt633 
         Caption         =   "&8.Enfermedades"
      End
   End
   Begin VB.Menu dlo333 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "thclinic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cli44_Click()
    tclinica.Show 1

End Sub

Private Sub Command1_Click()
    ttratame.Show 1

End Sub

Private Sub d7733_Click()
    ttipoafi.Show 1

End Sub

Private Sub djw221_Click()
    ttipocon.Show 1

End Sub

Private Sub dk8822_Click()
    ttipoaut.Show 1

End Sub

Private Sub dksef_Click()
    tseguro.Show 1

End Sub

Private Sub dlo333_Click()
    thclinic.Hide
    Unload thclinic

End Sub

Private Sub engt633_Click()
    tenferme.Show 1

End Sub

Private Sub image14_Click()
    tdiagnos.Show 1

End Sub

Private Sub Image2_Click()
    tconsult.Show 1

End Sub

Private Sub Image3_Click()
    tnclie.Show 1

End Sub

Private Sub Image5_Click()
    tasiste.Show 1

End Sub

Private Sub medi8844_Click()
    tmedico.Show 1

End Sub

Private Sub refr664_Click()
    treferen.Show 1

End Sub
