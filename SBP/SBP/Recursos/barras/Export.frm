VERSION 5.00
Begin VB.Form Export 
   Caption         =   "Export to Bitmap"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7080
   ControlBox      =   0   'False
   Icon            =   "Export.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4965
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Copy to Clipboard"
      Height          =   372
      Left            =   240
      TabIndex        =   4
      Top             =   4440
      Width           =   1932
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3120
      Top             =   2280
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Export"
      Default         =   -1  'True
      Height          =   372
      Left            =   4440
      TabIndex        =   3
      Top             =   4440
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   5760
      TabIndex        =   2
      Top             =   4440
      Width           =   1092
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H8000000C&
      Height          =   4092
      Left            =   0
      ScaleHeight     =   7.117
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   12.383
      TabIndex        =   0
      Top             =   0
      Width           =   7080
      Begin VB.PictureBox label 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1932
         Left            =   120
         ScaleHeight     =   3.413
         ScaleMode       =   7  'Centimeter
         ScaleWidth      =   4.683
         TabIndex        =   1
         Top             =   120
         Width           =   2652
      End
   End
End
Attribute VB_Name = "Export"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    
    Unload Me
    
End Sub

Function ResizeElements()

    Picture1.Height = Me.ScaleHeight - Command1.Height - 300
    Command1.Top = Me.ScaleHeight - Command1.Height - 160
    Command2.Top = Me.ScaleHeight - Command2.Height - 160
    Command3.Top = Me.ScaleHeight - Command2.Height - 160
    
    Command1.Left = Me.ScaleWidth - Command1.Width - 200
    Command2.Left = Me.ScaleWidth - Command1.Width - Command2.Width - 300

End Function

Private Sub Command2_Click()

    On Error GoTo Err:

    With tcxbarra.CommonDialog1
        
        .Filter = "Bitmap File (*.bmp)|*.bmp|"
        .CancelError = True
        .DialogTitle = "Export to Bitmap"
        .ShowSave
        Set label.Picture = label.Image
        SavePicture label.Picture, .FileName

    End With

    Unload Me

Err:
    Exit Sub

End Sub

Private Sub Command3_Click()

    Clipboard.Clear
    Set label.Picture = label.Image
    Clipboard.SetData label.Picture

End Sub

Private Sub Form_Load()

    Me.Caption = "Export to Bitmap"
    Set tcxbarra.LabelImage.Picture = tcxbarra.LabelImage.Image
    label.Width = tcxbarra.label.Width
    label.Height = tcxbarra.label.Height
    label.Picture = tcxbarra.LabelImage.Picture

End Sub

Private Sub Form_Resize()

    ResizeElements

End Sub

Private Sub Form_Unload(Cancel As Integer)

    tcxbarra.Enabled = True
    'tcxbarra.Show
    Unload Me

End Sub
