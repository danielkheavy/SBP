VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form InsertPicture 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Picture"
   ClientHeight    =   2910
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   5295
   ControlBox      =   0   'False
   Icon            =   "InsertPicture.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5295
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Apply"
      Default         =   -1  'True
      Height          =   372
      Left            =   2640
      TabIndex        =   5
      Top             =   1200
      Width           =   1092
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   3840
      TabIndex        =   4
      Top             =   1200
      Width           =   1212
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3480
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Preview"
      Height          =   1572
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1932
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   972
         Left            =   240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1452
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Picture"
      Height          =   852
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4932
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   252
         Left            =   4320
         TabIndex        =   2
         Top             =   360
         Width           =   372
      End
      Begin VB.TextBox Text1 
         Height          =   288
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   3972
      End
   End
   Begin VB.Label HiddenObjectIndex 
      Caption         =   "0"
      Height          =   252
      Left            =   2640
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   1692
   End
End
Attribute VB_Name = "InsertPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Item As ListItem

Private Sub Command1_Click()

    On Error GoTo Err:
    
    With CommonDialog1
        .Filter = "All picture files (*.jpg;*.gif;*.bmp)|*.jpg;*.gif;*.bmp|"
        .CancelError = True
        .ShowOpen
        Text1.Text = .FileName
        Image1.Picture = LoadPicture(.FileName)

    End With
    
Err:
    Exit Sub

End Sub

Private Sub Command2_Click()
    Unload Me

End Sub

Private Sub Command3_Click()
    
    If Text1.Text = "" Then
        MsgBox "Please select a picture!", vbCritical + vbOKOnly, "Insert Picture"
    Else

        If HiddenObjectIndex.Caption <> "0" Then

            'modify current picture
            With tcxbarra.ObjectImage(HiddenObjectIndex.Caption)
                .Picture = LoadPicture(Text1.Text)
                .ZOrder (0)
                .Visible = True

            End With

            With tcxbarra.ObjectImageURL(tcxbarra.ObjectImageURL.UBound)
                .Caption = Text1.Text

            End With

            tcxbarra.DisplayMovePosition.Visible = False
        Else
            'add new picture
            Load tcxbarra.ObjectImage(tcxbarra.ObjectImage.count)
            Load tcxbarra.ObjectImageURL(tcxbarra.ObjectImageURL.count)
            Load tcxbarra.ObjectImageList(tcxbarra.ObjectImageList.count)
            Set Item = tcxbarra.ListView1.ListItems.Add(, , "Picture " & tcxbarra.ObjectImage.UBound, , 2)
            tcxbarra.ListView1.ListItems(tcxbarra.ListView1.ListItems.count).Checked = True

            With tcxbarra.ObjectImage(tcxbarra.ObjectImage.UBound)
                .Picture = LoadPicture(Text1.Text)
                .ZOrder (0)
                .Left = 0.5
                .Top = 0.5
                .Visible = True

            End With

            With tcxbarra.ObjectImageURL(tcxbarra.ObjectImageURL.UBound)
                .Caption = Text1.Text

            End With

            With tcxbarra.ObjectImageList(tcxbarra.ObjectImageList.UBound)
                .Caption = tcxbarra.ListView1.ListItems.count

            End With

        End If

        Unload Me

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    tcxbarra.Enabled = True
    InsertPicture.Hide
    Unload Me

End Sub
