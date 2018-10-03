VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form tpropie 
   Caption         =   "Project Properties"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3855
   ControlBox      =   0   'False
   Icon            =   "tpropie.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5175
   ScaleWidth      =   3855
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   4560
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply"
      Default         =   -1  'True
      Height          =   372
      Left            =   1080
      TabIndex        =   10
      Top             =   4560
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2280
      TabIndex        =   9
      Top             =   4560
      Width           =   1212
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Label border"
      Height          =   252
      Left            =   240
      TabIndex        =   8
      Top             =   4080
      Value           =   1  'Checked
      Width           =   1452
   End
   Begin VB.Frame Frame2 
      Caption         =   "Background"
      Height          =   1452
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   3372
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFFFF&
         Height          =   372
         Left            =   240
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   12
         Top             =   840
         Width           =   252
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   852
         Left            =   540
         Picture         =   "tpropie.frx":0D4A
         ScaleHeight     =   795
         ScaleWidth      =   2595
         TabIndex        =   7
         Top             =   360
         Width           =   2652
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         Height          =   372
         Left            =   240
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   6
         Top             =   360
         Width           =   252
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Label size"
      Height          =   1932
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   3372
      Begin VB.ComboBox Combo1 
         Height          =   288
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   960
         Width           =   2052
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   252
         Left            =   960
         Max             =   3000
         Min             =   100
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Value           =   600
         Width           =   2052
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   252
         Left            =   960
         Max             =   5000
         Min             =   100
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   636
         Value           =   400
         Width           =   2052
      End
      Begin VB.Label Label4 
         Caption         =   "Preset:"
         Height          =   252
         Left            =   240
         TabIndex        =   13
         Top             =   1000
         Width           =   612
      End
      Begin VB.Label Label3 
         Caption         =   "6 x 4 cm"
         Height          =   252
         Left            =   960
         TabIndex        =   11
         Top             =   1440
         Width           =   1932
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Width:"
         Height          =   252
         Left            =   240
         TabIndex        =   4
         Top             =   380
         Width           =   492
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Height:"
         Height          =   252
         Left            =   240
         TabIndex        =   3
         Top             =   684
         Width           =   612
      End
   End
   Begin VB.Image Image1 
      Height          =   1020
      Left            =   4320
      Picture         =   "tpropie.frx":BBCC
      Top             =   3720
      Visible         =   0   'False
      Width           =   3270
   End
End
Attribute VB_Name = "tpropie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Click()
    
    If Combo1.Text = "Business Card" Then
        HScroll1.Value = "850"
        HScroll2.Value = "550"
    ElseIf Combo1.Text = "Greeting Card" Then
        HScroll1.Value = "850"
        HScroll2.Value = "1100"
    ElseIf Combo1.Text = "Post Card (5.5 x 4.25)" Then
        HScroll1.Value = "550"
        HScroll2.Value = "425"
    ElseIf Combo1.Text = "Post Card (8.5 x 5.47)" Then
        HScroll1.Value = "850"
        HScroll2.Value = "547"
    ElseIf Combo1.Text = "CD Booklet" Then
        HScroll1.Value = "1210"
        HScroll2.Value = "1210"
    ElseIf Combo1.Text = "CD Inlay" Then
        HScroll1.Value = "1500"
        HScroll2.Value = "1180"
    ElseIf Combo1.Text = "DVD Booklet" Then
        HScroll1.Value = "1210"
        HScroll2.Value = "1800"
    ElseIf Combo1.Text = "DVD Insert" Then
        HScroll1.Value = "2730"
        HScroll2.Value = "1800"

    End If

End Sub

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command2_Click()

    tcxbarra.label.BackColor = Picture2.BackColor

    If Check1.Value = 1 Then
        tcxbarra.label.BorderStyle = 1
    Else
        tcxbarra.label.BorderStyle = 0

    End If

    tcxbarra.label.Width = HScroll1.Value / 100
    tcxbarra.label.Height = HScroll2.Value / 100
    tcxbarra.LabelShadow.Width = HScroll1.Value / 100
    tcxbarra.LabelShadow.Height = HScroll2.Value / 100
    tcxbarra.StatusBar1.Panels(2).Text = "Label: " & Round(HScroll1.Value / 100, 2) & " x " & Round(HScroll2.Value / 100, 2) & " cm"
    
    UnloadGuidelines

    If Combo1.Text = "CD Inlay" Then
        InsertGuideline 1, 0.59, 0.59, 0, tcxbarra.label.ScaleHeight
        InsertGuideline 1, 14.4, 14.4, 0, tcxbarra.label.ScaleHeight

    End If

    If Combo1.Text = "DVD Insert" Then
        InsertGuideline 1, 13, 13, 0, tcxbarra.label.ScaleHeight
        InsertGuideline 1, 14.2, 14.2, 0, tcxbarra.label.ScaleHeight

    End If
    
    tcxbarra.Enabled = True
    
    'tcxbarra.Show
    
    ResizeElements
    tpropie.Hide
    Unload Me

End Sub

Private Sub Form_Load()

    Picture1.Height = Image1.Height

    Combo1.AddItem "Select..."
    Combo1.AddItem "Business Card"
    Combo1.AddItem "Greeting Card"
    Combo1.AddItem "Post Card (5.5 x 4.25)"
    Combo1.AddItem "Post Card (8.5 x 5.47)"
    Combo1.AddItem "CD Booklet"
    Combo1.AddItem "CD Inlay"
    Combo1.AddItem "DVD Booklet"
    Combo1.AddItem "DVD Insert"
    Combo1.Text = "Select..."

    Picture2.BackColor = tcxbarra.label.BackColor

    If tcxbarra.label.BorderStyle = 1 Then
        Check1.Value = 1
    Else
        Check1.Value = 0

    End If

    HScroll1.Value = tcxbarra.label.Width * 100
    HScroll2.Value = tcxbarra.label.Height * 100

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    tcxbarra.Enabled = True
    'tpropie.Hide
    Unload Me

End Sub

Private Sub HScroll1_Change()

    Label3.Caption = Round(HScroll1.Value / 100, 2) & " x " & Round(HScroll2.Value / 100, 2) & " cm"

End Sub

Private Sub HScroll1_Scroll()
    
    Label3.Caption = Round(HScroll1.Value / 100, 2) & " x " & Round(HScroll2.Value / 100, 2) & " cm"

End Sub

Private Sub HScroll2_Change()

    Label3.Caption = Round(HScroll1.Value / 100, 2) & " x " & Round(HScroll2.Value / 100, 2) & " cm"

End Sub

Private Sub HScroll2_Scroll()

    Label3.Caption = Round(HScroll1.Value / 100, 2) & " x " & Round(HScroll2.Value / 100, 2) & " cm"

End Sub

Private Sub Picture1_Click()

    Picture2.BackColor = Picture3.BackColor

End Sub

Private Sub Picture1_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

    Picture3.BackColor = Picture1.POINT(X, Y)

End Sub

Private Sub Picture2_Click()

    On Error GoTo Err:

    With CommonDialog1
        .color = Picture2.BackColor
        .Flags = 1
        .ShowColor
        Picture2.BackColor = .color

    End With

Err:
    Exit Sub

End Sub
