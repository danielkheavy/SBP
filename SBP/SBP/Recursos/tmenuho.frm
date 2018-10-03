VERSION 5.00
Begin VB.Form tmenuho 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hotel Control"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   13950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   13950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton image4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check-In"
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
      Left            =   8400
      MaskColor       =   &H8000000E&
      Picture         =   "tmenuho.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton image3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check-Out"
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
      Left            =   10200
      Picture         =   "tmenuho.frx":0576
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton image12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reservas"
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
      Left            =   6600
      Picture         =   "tmenuho.frx":0B5F
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   13890
      TabIndex        =   0
      Top             =   0
      Width           =   13950
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   7935
      Left            =   0
      Stretch         =   -1  'True
      Top             =   840
      Width           =   6495
   End
   Begin VB.Menu tabldre 
      Caption         =   "&Tablas"
      Begin VB.Menu salorr 
         Caption         =   "&1.Salones"
      End
      Begin VB.Menu habsalo 
         Caption         =   "&2.Habitaciones Salones"
      End
      Begin VB.Menu po99 
         Caption         =   "&3.Tipos Habitaciones"
      End
      Begin VB.Menu po922 
         Caption         =   "&4.Personal"
      End
      Begin VB.Menu cli89cli 
         Caption         =   "&5.Clientes"
      End
   End
   Begin VB.Menu reserva66 
      Caption         =   "&Reserva"
   End
   Begin VB.Menu cjurion4 
      Caption         =   "&CheckIn"
   End
   Begin VB.Menu fki84 
      Caption         =   "&CheckOut"
   End
   Begin VB.Menu fju744 
      Caption         =   "&Reportes"
   End
   Begin VB.Menu flo44 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tmenuho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub flo44_Click()
tmenuho.Hide
Unload tmenuho
End Sub

Private Sub image12_Click()
treserva.Show 1
End Sub
