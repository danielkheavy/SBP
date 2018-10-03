VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form InsertShape 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Shape"
   ClientHeight    =   3315
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   7470
   ControlBox      =   0   'False
   Icon            =   "InsertShape.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   7470
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   2880
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply"
      Default         =   -1  'True
      Height          =   372
      Left            =   4560
      TabIndex        =   14
      Top             =   2760
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   5880
      TabIndex        =   13
      Top             =   2760
      Width           =   1332
   End
   Begin VB.Frame Frame2 
      Caption         =   "Layout"
      Height          =   2412
      Left            =   4800
      TabIndex        =   7
      Top             =   120
      Width           =   2412
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         Height          =   252
         Left            =   1680
         ScaleHeight     =   195
         ScaleWidth      =   315
         TabIndex        =   12
         Top             =   1200
         Width           =   372
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   252
         Left            =   1680
         ScaleHeight     =   195
         ScaleWidth      =   315
         TabIndex        =   11
         Top             =   840
         Width           =   372
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Border"
         Height          =   252
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Value           =   1  'Checked
         Width           =   972
      End
      Begin VB.Label Label5 
         Caption         =   "Border color:"
         Height          =   252
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   972
      End
      Begin VB.Label Label4 
         Caption         =   "Background color:"
         Height          =   252
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1452
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sizes"
      Height          =   1812
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   4452
      Begin VB.CheckBox Check2 
         Caption         =   "Keep same sizes"
         Height          =   252
         Left            =   1800
         TabIndex        =   18
         Top             =   1320
         Width           =   1692
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   252
         Left            =   1800
         Max             =   1000
         Min             =   10
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   840
         Value           =   10
         Width           =   2412
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   252
         Left            =   1800
         Max             =   1000
         Min             =   10
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   480
         Value           =   10
         Width           =   2412
      End
      Begin VB.TextBox Text2 
         Height          =   288
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "1,00"
         Top             =   840
         Width           =   612
      End
      Begin VB.TextBox Text1 
         Height          =   288
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "1,00"
         Top             =   480
         Width           =   612
      End
      Begin VB.Label Label3 
         Caption         =   "Height:"
         Height          =   252
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   612
      End
      Begin VB.Label Label2 
         Caption         =   "Width:"
         Height          =   252
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   492
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1692
   End
   Begin VB.Label HiddenObjectIndex 
      Caption         =   "0"
      Height          =   252
      Left            =   1560
      TabIndex        =   15
      Top             =   3000
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label Label1 
      Caption         =   "Shape:"
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   612
   End
End
Attribute VB_Name = "InsertShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()

    If Combo1.Text = "Circle" Then
        HScroll2.Enabled = False
    Else
        HScroll2.Enabled = True

    End If

End Sub

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command2_Click()

    If Combo1.Text = "" Then
        MsgBox "No shape selected!", vbCritical + vbOKOnly, "Insert Shape"
    Else

        If HiddenObjectIndex.Caption <> 0 Then

            With tcxbarra.ObjectShape(HiddenObjectIndex.Caption)

                If Combo1.Text = "Rectangle" Then
                    .Width = Text1.Text
                    .Height = Text2.Text
                ElseIf Combo1.Text = "Circle" Then
                    .Width = Text1.Text
                    .Height = Text1.Text

                End If

            End With

            With tcxbarra.ObjectShapeForm(HiddenObjectIndex.Caption)

                If Combo1.Text = "Rectangle" Then
                    .Width = Text1.Text
                    .Height = Text2.Text
                    .Shape = 0
                ElseIf Combo1.Text = "Circle" Then
                    .Width = Text1.Text
                    .Height = Text1.Text
                    .Shape = 3

                End If

                If Check1.Value = 1 Then
                    .BorderStyle = 1
                Else
                    .BorderStyle = 0

                End If

                .BorderColor = Picture2.BackColor
                .FillColor = Picture1.BackColor

            End With

        Else
            'add new shape
            Load tcxbarra.ObjectShape(tcxbarra.ObjectShape.count)
            Load tcxbarra.ObjectShapeForm(tcxbarra.ObjectShapeForm.count)
            Load tcxbarra.ObjectShapeList(tcxbarra.ObjectShapeList.count)
            Set Item = tcxbarra.ListView1.ListItems.Add(, , "Shape " & tcxbarra.ObjectShape.UBound, , 3)
            tcxbarra.ListView1.ListItems(tcxbarra.ListView1.ListItems.count).Checked = True

            With tcxbarra.ObjectShape(tcxbarra.ObjectShape.UBound)
                .ZOrder (0)
                .Left = 0.5
                .Top = 0.5

                If Combo1.Text = "Rectangle" Then
                    .Width = Text1.Text
                    .Height = Text2.Text
                ElseIf Combo1.Text = "Circle" Then
                    .Width = Text1.Text
                    .Height = Text1.Text

                End If

                .Visible = True

            End With

            With tcxbarra.ObjectShapeForm(tcxbarra.ObjectShapeForm.UBound)
                .ZOrder (0)
                .Left = 0.5
                .Top = 0.5

                If Combo1.Text = "Rectangle" Then
                    .Width = Text1.Text
                    .Height = Text2.Text
                    .Shape = 0
                ElseIf Combo1.Text = "Circle" Then
                    .Width = Text1.Text
                    .Height = Text1.Text
                    .Shape = 3

                End If

                If Check1.Value = 1 Then
                    .BorderStyle = 1
                Else
                    .BorderStyle = 0

                End If

                .BorderColor = Picture2.BackColor
                .FillColor = Picture1.BackColor
                .Visible = True

            End With

            With tcxbarra.ObjectShapeList(tcxbarra.ObjectShapeList.UBound)
                .Caption = tcxbarra.ListView1.ListItems.count

            End With

        End If

        Unload Me

    End If

End Sub

Private Sub Form_Load()

    Combo1.AddItem "Rectangle"
    Combo1.AddItem "Circle"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    tcxbarra.Enabled = True
    InsertShape.Hide
    Unload Me

End Sub

Private Sub HScroll1_Change()

    If Check2.Value = 1 Then
        Text2.Text = Round(HScroll1.Value / 100, 2)
        HScroll2.Value = HScroll1.Value

    End If

    Text1.Text = Round(HScroll1.Value / 100, 2)

End Sub

Private Sub HScroll1_Scroll()

    If Check2.Value = 1 Then
        Text2.Text = Round(HScroll1.Value / 100, 2)
        HScroll2.Value = HScroll1.Value

    End If

    Text1.Text = Round(HScroll1.Value / 100, 2)

End Sub

Private Sub HScroll2_Change()

    If Check2.Value = 1 Then
        Text1.Text = Round(HScroll2.Value / 100, 2)
        HScroll1.Value = HScroll2.Value

    End If

    Text2.Text = Round(HScroll2.Value / 100, 2)

End Sub

Private Sub HScroll2_Scroll()

    If Check2.Value = 1 Then
        Text1.Text = Round(HScroll2.Value / 100, 2)
        HScroll1.Value = HScroll2.Value

    End If

    Text2.Text = Round(HScroll2.Value / 100, 2)

End Sub

Private Sub Picture1_Click()

    On Error GoTo Err:

    With CommonDialog1
        .color = Picture1.BackColor
        .CancelError = True
        .Flags = 1
        .ShowColor
        Picture1.BackColor = .color

    End With

Err:
    Exit Sub

End Sub

Private Sub Picture2_Click()

    On Error GoTo Err:

    With CommonDialog1
        .color = Picture2.BackColor
        .CancelError = True
        .Flags = 1
        .ShowColor
        Picture2.BackColor = .color

    End With

Err:
    Exit Sub

End Sub
