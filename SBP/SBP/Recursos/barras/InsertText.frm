VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form InsertText 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Text"
   ClientHeight    =   3900
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   5580
   ControlBox      =   0   'False
   Icon            =   "InsertText.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5580
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "Preview"
      Height          =   1572
      Left            =   3000
      TabIndex        =   17
      Top             =   2160
      Width           =   2412
      Begin VB.Label Label4 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AaBb"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   972
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1932
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Color"
      Height          =   852
      Left            =   3000
      TabIndex        =   13
      Top             =   1200
      Width           =   2412
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         Height          =   252
         Left            =   1440
         ScaleHeight     =   195
         ScaleWidth      =   315
         TabIndex        =   14
         Top             =   360
         Width           =   372
      End
      Begin VB.Label Label3 
         Caption         =   "Font color:"
         Height          =   252
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   852
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Font"
      Height          =   1932
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2772
      Begin VB.CheckBox Check4 
         Caption         =   "Strikeout"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   12
         Top             =   1440
         Width           =   972
      End
      Begin VB.ComboBox Combo2 
         Height          =   288
         Left            =   720
         TabIndex        =   10
         Text            =   "12"
         Top             =   720
         Width           =   732
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Bold"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   732
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Italic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   732
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Underline"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   7
         Top             =   1200
         Width           =   1092
      End
      Begin VB.ComboBox Combo1 
         Height          =   288
         Left            =   720
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   1812
      End
      Begin VB.Label Label2 
         Caption         =   "Size:"
         Height          =   252
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   372
      End
      Begin VB.Label Label1 
         Caption         =   "Font:"
         Height          =   252
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   492
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4320
      Top             =   4320
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply"
      Default         =   -1  'True
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   1320
      TabIndex        =   2
      Top             =   3360
      Width           =   1092
   End
   Begin VB.Frame Frame1 
      Caption         =   "Caption"
      Height          =   852
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5292
      Begin VB.TextBox Text1 
         Height          =   288
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   4812
      End
   End
   Begin VB.Label HiddenObjectIndex 
      Caption         =   "0"
      Height          =   252
      Left            =   4800
      TabIndex        =   16
      Top             =   4320
      Visible         =   0   'False
      Width           =   372
   End
End
Attribute VB_Name = "InsertText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim I    As Integer

Dim Item As ListItem

Private Sub Combo1_Click()
    Label4.FontName = Combo1.Text

End Sub

Private Sub Combo2_Click()
    Label4.FontSize = Combo2.Text

End Sub

Private Sub Command1_Click()
    Unload Me

End Sub

Private Sub Command2_Click()
    
    If Text1.Text <> "" And Combo1.Text <> "" Then
        
        If HiddenObjectIndex.Caption <> 0 Then

            With tcxbarra.ObjectText(HiddenObjectIndex.Caption)
                .Caption = Text1.Text
                .FontName = Combo1.Text
                .FontSize = Combo2.Text
                .ForeColor = Picture1.BackColor

                If Check1.Value = 1 Then
                    .FontBold = True
                Else
                    .FontBold = False

                End If

                If Check2.Value = 1 Then
                    .FontItalic = True
                Else
                    .FontItalic = False

                End If

                If Check3.Value = 1 Then
                    .FontUnderline = True
                Else
                    .FontUnderline = False

                End If

                If Check4.Value = 1 Then
                    .FontStrikethru = True
                Else
                    .FontStrikethru = False

                End If

                .ZOrder (0)
                .Visible = True
                tcxbarra.DisplayMovePosition.Visible = False

            End With

        Else
            Load tcxbarra.ObjectText(tcxbarra.ObjectText.count)
            Load tcxbarra.ObjectTextList(tcxbarra.ObjectTextList.count)
            Set Item = tcxbarra.ListView1.ListItems.Add(, , "Text " & tcxbarra.ObjectText.UBound, , 1)
            tcxbarra.ListView1.ListItems(tcxbarra.ListView1.ListItems.count).Checked = True

            With tcxbarra.ObjectText(tcxbarra.ObjectText.UBound)
                .Caption = Text1.Text
                .FontName = Combo1.Text
                .FontSize = Combo2.Text
                .ForeColor = Picture1.BackColor

                If Check1.Value = 1 Then
                    .FontBold = True

                End If

                If Check2.Value = 1 Then
                    .FontItalic = True

                End If

                If Check3.Value = 1 Then
                    .FontUnderline = True

                End If

                If Check4.Value = 1 Then
                    .FontStrikethru = True

                End If

                .ZOrder (0)
                .Left = 0.5
                .Top = 0.5
                .Visible = True

            End With

            With tcxbarra.ObjectTextList(tcxbarra.ObjectTextList.UBound)
                .Caption = tcxbarra.ListView1.ListItems.count

            End With

        End If
        
        Unload Me
    
    Else
        MsgBox "Please fill in a text and fontname!", vbCritical + vbOKOnly, "Insert Text"

    End If

End Sub

Private Sub Form_Load()
    
    For I = 0 To Screen.FontCount - 1
        Combo1.AddItem Screen.Fonts(I)
    Next I

    For I = 6 To 24
        Combo2.AddItem I
    Next I
    
    If HiddenObjectIndex.Caption <> 0 Then
        Combo1.Text = tcxbarra.ObjectText(HiddenObjectIndex.Caption).FontName
        Combo2.Text = tcxbarra.ObjectText(HiddenObjectIndex.Caption).FontSize
        Check1.Value = tcxbarra.ObjectText(HiddenObjectIndex.Caption).FontBold
        Check2.Value = tcxbarra.ObjectText(HiddenObjectIndex.Caption).FontItalic
        Check3.Value = tcxbarra.ObjectText(HiddenObjectIndex.Caption).FontUnderline
        Check4.Value = tcxbarra.ObjectText(HiddenObjectIndex.Caption).FontStrikethru
        Picture1.BackColor = tcxbarra.ObjectText(HiddenObjectIndex.Caption).ForeColor
    Else
        Combo1.Text = "Arial"
        Combo2.Text = "11"

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'tcxbarra.Enabled = True
    InsertText.Hide
    Unload InsertText

End Sub

Private Sub Picture1_Click()
    
    On Error GoTo Err:
    
    With CommonDialog1
        .CancelError = True
        .color = Picture1.BackColor
        .ShowColor
        .Flags = 1
        Picture1.BackColor = .color

    End With

Err:
    Exit Sub

End Sub

Private Sub Picture2_Click()

    On Error GoTo Err:

    With CommonDialog1
        .color = Picture2.BackColor
        .ShowColor
        .Flags = 1
        .CancelError = 1
        Picture2.BackColor = .color

    End With

Err:
    Exit Sub

End Sub
