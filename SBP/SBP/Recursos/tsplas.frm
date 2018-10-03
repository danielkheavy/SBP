VERSION 5.00
Begin VB.Form tsplas 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Height          =   3795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7485
      Begin VB.PictureBox Picture2 
         Height          =   1005
         Left            =   6000
         Picture         =   "tsplas.frx":0000
         ScaleHeight     =   945
         ScaleWidth      =   1185
         TabIndex        =   6
         Top             =   240
         Width           =   1245
      End
      Begin VB.PictureBox Picture1 
         Height          =   1005
         Left            =   90
         Picture         =   "tsplas.frx":0BA3
         ScaleHeight     =   945
         ScaleWidth      =   1185
         TabIndex        =   1
         Top             =   270
         Width           =   1245
      End
      Begin VB.Timer Timer1 
         Interval        =   2000
         Left            =   180
         Top             =   2610
      End
      Begin VB.Label lblCompanyProduct 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "Bienvenido a su Maravilloso herramienta de Trabajo"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   885
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   4425
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         Caption         =   "SISTEMA CALIPSO"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   960
         TabIndex        =   4
         Top             =   1440
         Width           =   5955
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         Caption         =   "Version :  5.1.0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5070
         TabIndex        =   3
         Top             =   2250
         Width           =   1515
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "Derechos Reservados(R)  NOEL  YNNHOJ  - DERILAK.MOC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   465
         Left            =   120
         TabIndex        =   2
         Top             =   3120
         Width           =   7365
      End
   End
End
Attribute VB_Name = "tsplas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub


Private Sub Timer1_Timer()
menup.Show 1
Unload tsplas
End Sub

