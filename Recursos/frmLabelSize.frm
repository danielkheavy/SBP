VERSION 5.00
Begin VB.Form frmLabelSize 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar Tamano etiqueta"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLabelSize.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4710
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtWidth 
      Height          =   345
      Left            =   1680
      TabIndex        =   9
      Top             =   1560
      Width           =   1155
   End
   Begin VB.TextBox txtHeight 
      Height          =   345
      Left            =   1680
      TabIndex        =   8
      Top             =   1140
      Width           =   1155
   End
   Begin VB.Frame fraScale 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Scale:"
      Height          =   675
      Left            =   180
      TabIndex        =   4
      Top             =   180
      Width           =   4335
      Begin VB.OptionButton optScale 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centimeter"
         Height          =   240
         Index           =   2
         Left            =   2880
         TabIndex        =   7
         Top             =   300
         Width           =   1335
      End
      Begin VB.OptionButton optScale 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Inch"
         Height          =   240
         Index           =   1
         Left            =   1500
         TabIndex        =   6
         Top             =   300
         Width           =   1335
      End
      Begin VB.OptionButton optScale 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pixel"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdDone 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Set"
      Height          =   435
      Index           =   1
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Transfer the Item."
      Top             =   2460
      Width           =   975
   End
   Begin VB.CommandButton cmdDone 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancel"
      Height          =   435
      Index           =   0
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Transfer the Item."
      Top             =   2460
      Width           =   975
   End
   Begin VB.Label lblWidth 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Scale Width:"
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   1620
      Width           =   1335
   End
   Begin VB.Label lblHeight 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Scale Height:"
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "frmLabelSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bCancel As Boolean

Public Property Get Cancelled() As Boolean
    Cancelled = m_bCancel

End Property

Private Sub cmdDone_Click(Index As Integer)

    Select Case Index

        Case 0  'Cancel
            m_bCancel = True
            Me.Hide
            
        Case 1  'Set
            m_bCancel = False
            
    End Select
    
End Sub

Private Sub iuu8_Click()

End Sub
