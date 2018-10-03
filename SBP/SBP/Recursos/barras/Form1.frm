VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form tcxbarra
   Caption         =   "Form1"
   ClientHeight    =   7710
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11580
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   11580
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   3960
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   16777215
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0D4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":124C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":174E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox LayerBar 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7485
      Left            =   0
      ScaleHeight     =   7485
      ScaleWidth      =   3000
      TabIndex        =   5
      Top             =   0
      Width           =   3000
      Begin VB.Frame Frame3 
         BackColor       =   &H80000004&
         Caption         =   "Object browser"
         Height          =   4812
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2772
         Begin MSComctlLib.ListView ListView1 
            Height          =   3252
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   2532
            _ExtentX        =   4471
            _ExtentY        =   5741
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            Icons           =   "ImageList1"
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   3528
            EndProperty
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   1920
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   225
      Left            =   0
      TabIndex        =   3
      Top             =   7485
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   397
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11906
            Text            =   "Untitled"
            TextSave        =   "Untitled"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Label: 6 x 4 cm"
            TextSave        =   "Label: 6 x 4 cm"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2831
            MinWidth        =   2822
            Text            =   "Position: 0 x 0"
            TextSave        =   "Position: 0 x 0"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Object: 0 x 0"
            TextSave        =   "Object: 0 x 0"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox WorkBar 
      Align           =   3  'Align Left
      BackColor       =   &H8000000C&
      Height          =   7485
      Left            =   3000
      ScaleHeight     =   13.097
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   10.292
      TabIndex        =   0
      Top             =   0
      Width           =   5892
      Begin VB.VScrollBar VScroll1 
         Enabled         =   0   'False
         Height          =   7092
         Left            =   5640
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   0
         Width           =   252
      End
      Begin VB.HScrollBar HScroll1 
         Enabled         =   0   'False
         Height          =   252
         Left            =   0
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   7200
         Width           =   5532
      End
      Begin VB.PictureBox Label 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5520
         Left            =   120
         ScaleHeight     =   9.684
         ScaleMode       =   7  'Centimeter
         ScaleWidth      =   8.44
         TabIndex        =   1
         Top             =   720
         Width           =   4819
         Begin VB.Label ObjectShapeList 
            BackStyle       =   0  'Transparent
            Caption         =   "{shape list}"
            Height          =   252
            Index           =   0
            Left            =   360
            TabIndex        =   29
            Top             =   4800
            Visible         =   0   'False
            Width           =   852
         End
         Begin VB.Label ObjectShape 
            BackStyle       =   0  'Transparent
            Height          =   372
            Index           =   0
            Left            =   360
            TabIndex        =   28
            Top             =   4800
            Width           =   612
         End
         Begin VB.Shape ObjectShapeForm 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   612
            Index           =   0
            Left            =   240
            Top             =   4680
            Visible         =   0   'False
            Width           =   972
         End
         Begin VB.Line HorizontalGuideline 
            BorderColor     =   &H00C0C0FF&
            Index           =   0
            Visible         =   0   'False
            X1              =   4.022
            X2              =   6.985
            Y1              =   4.657
            Y2              =   4.657
         End
         Begin VB.Line VerticalGuideline 
            BorderColor     =   &H00C0C0FF&
            Index           =   0
            Visible         =   0   'False
            X1              =   7.197
            X2              =   7.197
            Y1              =   0.847
            Y2              =   4.022
         End
         Begin VB.Shape DisplayObjectShape 
            BorderStyle     =   3  'Dot
            Height          =   372
            Left            =   2760
            Top             =   1560
            Visible         =   0   'False
            Width           =   372
         End
         Begin VB.Label BarNumberWidth 
            AutoSize        =   -1  'True
            Caption         =   "BarNumberWidth"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   4.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   120
            Left            =   240
            TabIndex        =   16
            Top             =   1680
            Visible         =   0   'False
            Width           =   816
         End
         Begin VB.Label ObjectImageList 
            BackStyle       =   0  'Transparent
            Caption         =   "{image list}"
            Height          =   252
            Index           =   0
            Left            =   1440
            TabIndex        =   13
            Top             =   600
            Visible         =   0   'False
            Width           =   1092
         End
         Begin VB.Label ObjectTextList 
            BackStyle       =   0  'Transparent
            Caption         =   "{text list}"
            Height          =   252
            Index           =   0
            Left            =   1560
            TabIndex        =   12
            Top             =   120
            Visible         =   0   'False
            Width           =   852
         End
         Begin VB.Line HorizontalCenterLine 
            BorderColor     =   &H00C0C0C0&
            Visible         =   0   'False
            X1              =   1.905
            X2              =   3.81
            Y1              =   2.752
            Y2              =   2.752
         End
         Begin VB.Line VerticalCenterLine 
            BorderColor     =   &H00C0C0C0&
            Visible         =   0   'False
            X1              =   3.81
            X2              =   3.81
            Y1              =   0.423
            Y2              =   2.328
         End
         Begin VB.Label ObjectImageURL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "{image object}"
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   8
            Top             =   600
            Visible         =   0   'False
            Width           =   1056
         End
         Begin VB.Line VerticalSplit 
            BorderColor     =   &H00FF8080&
            BorderStyle     =   3  'Dot
            Visible         =   0   'False
            X1              =   4.657
            X2              =   4.657
            Y1              =   1.058
            Y2              =   3.387
         End
         Begin VB.Line HorizontalSplit 
            BorderColor     =   &H00FF8080&
            BorderStyle     =   3  'Dot
            Visible         =   0   'False
            X1              =   0.635
            X2              =   5.503
            Y1              =   2.54
            Y2              =   2.54
         End
         Begin VB.Label DisplayMovePosition 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            Caption         =   "0 x 0"
            Height          =   156
            Left            =   1200
            TabIndex        =   7
            Top             =   1680
            Visible         =   0   'False
            Width           =   276
         End
         Begin VB.Image ObjectImage 
            Appearance      =   0  'Flat
            Height          =   372
            Index           =   0
            Left            =   120
            Top             =   600
            Visible         =   0   'False
            Width           =   852
         End
         Begin VB.Label ObjectText 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "{text object}"
            Height          =   192
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   120
            Visible         =   0   'False
            Width           =   840
         End
      End
      Begin VB.PictureBox LabelShadow 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5520
         Left            =   240
         ScaleHeight     =   9.737
         ScaleMode       =   7  'Centimeter
         ScaleWidth      =   8.493
         TabIndex        =   2
         Top             =   840
         Width           =   4819
      End
      Begin VB.PictureBox BarTop 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   4.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   198
         Left            =   360
         ScaleHeight     =   0.344
         ScaleMode       =   7  'Centimeter
         ScaleWidth      =   0.45
         TabIndex        =   14
         Top             =   6720
         Width           =   252
         Begin VB.Line BarTopLine 
            BorderColor     =   &H00C00000&
            Visible         =   0   'False
            X1              =   0.212
            X2              =   0.212
            Y1              =   0
            Y2              =   0.212
         End
      End
      Begin VB.PictureBox BarLeft 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   4.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   120
         ScaleHeight     =   0.45
         ScaleMode       =   7  'Centimeter
         ScaleWidth      =   0.344
         TabIndex        =   15
         Top             =   6720
         Width           =   198
         Begin VB.Line BarLeftLine 
            BorderColor     =   &H00C00000&
            Visible         =   0   'False
            X1              =   0
            X2              =   0.212
            Y1              =   0.212
            Y2              =   0.212
         End
      End
      Begin VB.PictureBox BarEmpty 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   1560
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   17
         Top             =   3480
         Width           =   372
      End
      Begin VB.PictureBox LabelImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1932
         Left            =   360
         ScaleHeight     =   3.413
         ScaleMode       =   7  'Centimeter
         ScaleWidth      =   5.318
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   3012
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   4680
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   27
         Top             =   2520
         Width           =   252
      End
      Begin VB.Label HiddenY 
         Caption         =   "0"
         Height          =   252
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label HiddenX 
         Caption         =   "0"
         Height          =   252
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   1212
      End
   End
   Begin VB.Label columns 
      Caption         =   "1"
      Height          =   252
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   1332
   End
   Begin VB.Label rows 
      Caption         =   "1"
      Height          =   252
      Left            =   0
      TabIndex        =   23
      Top             =   240
      Width           =   1332
   End
   Begin VB.Label columnspacing 
      Caption         =   "1"
      Height          =   252
      Left            =   0
      TabIndex        =   22
      Top             =   480
      Width           =   1332
   End
   Begin VB.Label rowspacing 
      Caption         =   "1"
      Height          =   252
      Left            =   0
      TabIndex        =   21
      Top             =   720
      Width           =   1332
   End
   Begin VB.Label pagemargintop 
      Caption         =   "10"
      Height          =   252
      Left            =   1560
      TabIndex        =   20
      Top             =   0
      Width           =   1332
   End
   Begin VB.Label pagemarginleft 
      Caption         =   "10"
      Height          =   252
      Left            =   1560
      TabIndex        =   19
      Top             =   360
      Width           =   1332
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFile1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExport 
         Caption         =   "&Export to Bitmap"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuFile5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewMouseTracker 
         Caption         =   "Mouse tracker"
      End
      Begin VB.Menu mnuViewCenterLines 
         Caption         =   "Center lines"
      End
      Begin VB.Menu mnuView1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSizeLabel 
         Caption         =   "Display Size label"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditMove 
         Caption         =   "Move object"
         Enabled         =   0   'False
         Begin VB.Menu mnuEditUp 
            Caption         =   "Up"
         End
         Begin VB.Menu mnuEditDown 
            Caption         =   "Down"
         End
         Begin VB.Menu mnuEditLeft 
            Caption         =   "Left"
         End
         Begin VB.Menu mnuEditRight 
            Caption         =   "Right"
         End
         Begin VB.Menu mnuEditMove1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEditMoveMoveBy 
            Caption         =   "Move by..."
         End
      End
      Begin VB.Menu mnuEdit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditBringtofront 
         Caption         =   "Bring to front"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditSendtoback 
         Caption         =   "Send to back"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEdit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditUnselect 
         Caption         =   "Unselect"
         Enabled         =   0   'False
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "&Modify"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Project"
      Begin VB.Menu mnuInsertText 
         Caption         =   "Add &Text"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuInsertPicture 
         Caption         =   "Add &Picture"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuInserShape 
         Caption         =   "Add &Shape"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuInsert1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertProperties 
         Caption         =   "Project Prop&erties..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "tcxbarra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'define variables
Dim i As Integer
Dim VarObjectType As String
Dim VarObjectIndex As Integer
Dim VarListIndex As Integer
Dim MoveObject
Dim ObjectIsSelected
Dim SetCurrentX As String
Dim FileContent As String
Dim NewFileContent As String
Dim SetCurrentY As String
Dim TwipSize As String

Dim ArrayLineContent As Variant
Dim SourceFile As String

Private Sub label_KeyDown(KeyCode As Integer, Shift As Integer)

    If ObjectIsSelected = True Then
        If KeyCode = vbKeyLeft Then
            MoveObjectLeft
        ElseIf KeyCode = vbKeyRight Then
            MoveObjectRight
        ElseIf KeyCode = vbKeyUp Then
            MoveObjectUp
        ElseIf KeyCode = vbKeyDown Then
            MoveObjectDown
        End If
    End If

End Sub

Private Sub Label_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    StatusBar1.Panels(3).Text = "Position: " & Round(X, 2) & " x " & Round(Y, 2)
    HiddenX = X
    HiddenY = Y

    'mouse tracker
    If mnuViewMouseTracker.Checked = True Then
        HorizontalSplit.Visible = True
        HorizontalSplit.X1 = 0
        HorizontalSplit.X2 = label.Width
        HorizontalSplit.Y1 = Y
        HorizontalSplit.Y2 = Y
        HorizontalSplit.ZOrder (0)
        VerticalSplit.Visible = True
        VerticalSplit.X1 = X
        VerticalSplit.X2 = X
        VerticalSplit.Y1 = 0
        VerticalSplit.Y2 = label.Height
        VerticalSplit.ZOrder (0)
        HorizontalCenterLine.ZOrder (0)
        VerticalCenterLine.ZOrder (0)
    End If
    
    
    'move barlines
    Form1.BarTopLine.Visible = True
    Form1.BarLeftLine.Visible = True
    Form1.BarTopLine.X1 = X
    Form1.BarTopLine.X2 = X
    Form1.BarTopLine.Y1 = 0
    Form1.BarTopLine.Y2 = BarTop.Height
    Form1.BarLeftLine.X1 = 0
    Form1.BarLeftLine.X2 = BarLeft.Width
    Form1.BarLeftLine.Y1 = Y
    Form1.BarLeftLine.Y2 = Y

End Sub

Private Sub Label_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        PopupMenu mnuInsert
    End If

End Sub

Private Sub ListView1_Click()

    Dim ArrayObjectText As Variant

    If ListView1.ListItems.Count <> 0 Then
        ResetObjects
        ArrayObjectText = Split(ListView1.SelectedItem.Text, " ")
        If ArrayObjectText(0) = "Text" Then
            VarObjectType = "text"
        ElseIf ArrayObjectText(0) = "Picture" Then
            VarObjectType = "picture"
        ElseIf ArrayObjectText(0) = "Shape" Then
            VarObjectType = "shape"
        End If
        VarObjectIndex = ArrayObjectText(1)
        SelectObject
    End If

End Sub

Private Sub mnuEditBringtofront_Click()

    'Bring object to front
    If VarObjectType = "text" Then
        ObjectText(VarObjectIndex).ZOrder (0)
    ElseIf VarObjectType = "picture" Then
        ObjectImage(VarObjectIndex).ZOrder (0)
    ElseIf VarObjectType = "shape" Then
        ObjectShape(VarObjectIndex).ZOrder (0)
        ObjectShapeForm(VarObjectIndex).ZOrder (0)
    End If

End Sub

Private Sub mnuEditDelete_Click()

    If MsgBox("Are you sure you whant to delete this object?", vbQuestion + vbYesNo, "Delete object") = vbYes Then
        
        'delete selected object
        If VarObjectType = "text" Then
            ObjectText(VarObjectIndex).Visible = False
        ElseIf VarObjectType = "picture" Then
            ObjectImage(VarObjectIndex).Visible = False
            ObjectImageURL(VarObjectIndex).Visible = False
        ElseIf VarObjectType = "shape" Then
            ObjectShape(VarObjectIndex).Visible = False
            ObjectShapeForm(VarObjectIndex).Visible = False
        End If
        
        ListView1.ListItems.Remove (VarListIndex)
        ResetObjects
    
    End If

End Sub

Private Sub mnuEditModify_Click()

    'modify selected object
    If VarObjectType = "text" Then
        InsertText.Show
        InsertText.HiddenObjectIndex.Caption = VarObjectIndex
        InsertText.Text1.Text = ObjectText(VarObjectIndex).Caption
        InsertText.Combo1.Text = ObjectText(VarObjectIndex).FontName
        InsertText.Combo2.Text = ObjectText(VarObjectIndex).FontSize
        If ObjectText(VarObjectIndex).FontBold = True Then
            InsertText.Check1.Value = 1
        End If
        If ObjectText(VarObjectIndex).FontItalic = True Then
            InsertText.Check2.Value = 1
        End If
        If ObjectText(VarObjectIndex).FontUnderline = True Then
            InsertText.Check3.Value = 1
        End If
        If ObjectText(VarObjectIndex).FontStrikethru = True Then
            InsertText.Check4.Value = 1
        End If
        InsertText.Picture1.BackColor = ObjectText(VarObjectIndex).ForeColor
    ElseIf VarObjectType = "picture" Then
        InsertPicture.Show
        InsertPicture.HiddenObjectIndex.Caption = VarObjectIndex
        InsertPicture.Text1.Text = ObjectImageURL(VarObjectIndex).Caption
        InsertPicture.Image1.Picture = ObjectImage(VarObjectIndex).Picture
    ElseIf VarObjectType = "shape" Then
        InsertShape.HiddenObjectIndex.Caption = VarObjectIndex
        If ObjectShapeForm(VarObjectIndex).Shape = 0 Then
            InsertShape.Combo1.Text = "Rectangle"
        ElseIf ObjectShapeForm(VarObjectIndex).Shape = 3 Then
            InsertShape.Combo1.Text = "Circle"
        End If
        InsertShape.Show
        InsertShape.Text1.Text = ObjectShapeForm(VarObjectIndex).Width
        InsertShape.Text2.Text = ObjectShapeForm(VarObjectIndex).Height
        InsertShape.HScroll1.Value = Round(ObjectShapeForm(VarObjectIndex).Width * 100)
        InsertShape.HScroll2.Value = Round(ObjectShapeForm(VarObjectIndex).Height * 100)
        If ObjectShapeForm(VarObjectIndex).BorderStyle = 1 Then
            InsertShape.Check1.Value = 1
        Else
            InsertShape.Check1.Value = 0
        End If
        InsertShape.Picture1.BackColor = ObjectShapeForm(VarObjectIndex).FillColor
        InsertShape.Picture2.BackColor = ObjectShapeForm(VarObjectIndex).BorderColor
    End If
    
    Me.Enabled = False

End Sub

Private Sub mnuEditMoveMoveBy_Click()

    Dim NewTwipSize As String
    NewTwipSize = InputBox("Move objects with (cm):", "Move By", TwipSize)
    If IsNumeric(NewTwipSize) Then
        TwipSize = NewTwipSize
    End If

End Sub

Private Sub mnuEditSendtoback_Click()

    'Send object to back
    If VarObjectType = "text" Then
        ObjectText(VarObjectIndex).ZOrder (1)
    ElseIf VarObjectType = "picture" Then
        ObjectImage(VarObjectIndex).ZOrder (1)
    ElseIf VarObjectType = "shape" Then
        ObjectShape(VarObjectIndex).ZOrder (1)
        ObjectShapeForm(VarObjectIndex).ZOrder (1)
    End If

End Sub

Private Sub Form_Load()

    Me.Caption = App.Title
    mnuHelpAbout.Caption = "&About " & App.Title & "..."

    TwipSize = 0.1
    label.Width = 8.5
    label.Height = 5.5
    LabelShadow.Width = 8.5
    LabelShadow.Height = 5.5

    Me.Width = Screen.Width / 1.4
    Me.Height = Screen.Height / 1.4

    Me.Left = (Screen.Width - Me.ScaleWidth) / 2
    Me.Top = (Screen.Height - Me.ScaleHeight) / 2

    'set new shortcuts to menu
    mnuEditUp.Caption = "Up" & vbTab & "Up"
    mnuEditDown.Caption = "Down" & vbTab & "Down"
    mnuEditLeft.Caption = "Left" & vbTab & "Left"
    mnuEditRight.Caption = "Right" & vbTab & "Right"

End Sub

Private Sub Form_Initialize()
  InitControlsXP
End Sub


Function DrawPicture()

    LabelImage.Picture = LoadPicture("")

    'set label properties
    LabelImage.ScaleWidth = label.ScaleWidth
    LabelImage.ScaleHeight = label.ScaleHeight
    LabelImage.Width = label.Width
    LabelImage.Height = label.Height
    LabelImage.BackColor = label.BackColor

    'draw picture
    For i = 1 To ObjectImage.UBound
        If ObjectImage(i).Visible = True Then
            LabelImage.PaintPicture ObjectImage(i).Picture, ObjectImage(i).Left, ObjectImage(i).Top
        End If
    Next i

    'draw shape
    For i = 1 To ObjectShapeForm.UBound
        If ObjectShapeForm(i).Visible = True Then
            'draw rectangle
            If ObjectShapeForm(i).Shape = 0 Then
                'draw background color
                LabelImage.Line (ObjectShapeForm(i).Left, ObjectShapeForm(i).Top)-Step(ObjectShapeForm(i).Width, ObjectShapeForm(i).Height), ObjectShapeForm(i).FillColor, BF
                'draw border
                If ObjectShapeForm(i).BorderStyle = 1 Then
                    LabelImage.Line (ObjectShapeForm(i).Left, ObjectShapeForm(i).Top)-Step(ObjectShapeForm(i).Width, ObjectShapeForm(i).Height), ObjectShapeForm(i).BorderColor, B
                End If
            'draw rectangle
            ElseIf ObjectShapeForm(i).Shape = 3 Then
                'draw background color
                LabelImage.FillColor = ObjectShapeForm(i).FillColor
                LabelImage.FillStyle = vbFSSolid
                LabelImage.Circle (ObjectShapeForm(i).Left + (ObjectShapeForm(i).Width / 2), ObjectShapeForm(i).Top + (ObjectShapeForm(i).Height / 2)), ObjectShapeForm(i).Height / 2, ObjectShapeForm(i).BorderColor
                LabelImage.FillStyle = vbFSTransparent
            End If
        End If
    Next i

    'draw text
    For i = 1 To ObjectText.UBound
        If ObjectText(i).Visible = True Then
            LabelImage.CurrentX = ObjectText(i).Left
            LabelImage.CurrentY = ObjectText(i).Top
            LabelImage.FontName = ObjectText(i).FontName
            LabelImage.FontSize = ObjectText(i).FontSize
            LabelImage.FontBold = ObjectText(i).FontBold
            LabelImage.FontItalic = ObjectText(i).FontItalic
            LabelImage.FontUnderline = ObjectText(i).FontUnderline
            LabelImage.FontStrikethru = ObjectText(i).FontStrikethru
            LabelImage.ForeColor = ObjectText(i).ForeColor
            LabelImage.Print ObjectText(i).Caption
        End If
    Next i
    
    'draw border
    If label.BorderStyle = 1 Then
        LabelImage.Line (LabelImage.Width - 0.06, 0)-(0, LabelImage.Height - 0.06), 1, B
    End If

End Function

Private Sub Form_Resize()

    ResizeElements

End Sub

Function ResetObjects()
    
    'reset selections
    mnuEditMove.Enabled = False
    mnuEditBringtofront.Enabled = False
    mnuEditSendtoback.Enabled = False
    mnuEditUnselect.Enabled = False
    mnuEditDelete.Enabled = False
    mnuEditModify.Enabled = False
    
    'reset listview display
    For i = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(i).Bold = False
    Next i
    
    StatusBar1.Panels(4).Text = "Object: 0 x 0"
    DisplayMovePosition.Visible = False
    DisplayObjectShape.Visible = False
    
    ObjectIsSelected = False

End Function

Private Sub Label_Click()
    ResetObjects
End Sub

Function MoveObjectDown()

    'move selected object down
    If VarObjectType = "text" Then
        ObjectText(VarObjectIndex).Top = ObjectText(VarObjectIndex).Top + TwipSize
        DisplayPosition ObjectText(VarObjectIndex).Top, ObjectText(VarObjectIndex).Left
        DisplayShape ObjectText(VarObjectIndex).Top, ObjectText(VarObjectIndex).Left, ObjectText(VarObjectIndex).Width, ObjectText(VarObjectIndex).Height
    ElseIf VarObjectType = "picture" Then
        ObjectImage(VarObjectIndex).Top = ObjectImage(VarObjectIndex).Top + TwipSize
        DisplayPosition ObjectImage(VarObjectIndex).Top, ObjectImage(VarObjectIndex).Left
        DisplayShape ObjectImage(VarObjectIndex).Top, ObjectImage(VarObjectIndex).Left, ObjectImage(VarObjectIndex).Width, ObjectImage(VarObjectIndex).Height
    ElseIf VarObjectType = "shape" Then
        ObjectShape(VarObjectIndex).Top = ObjectShape(VarObjectIndex).Top + TwipSize
        ObjectShapeForm(VarObjectIndex).Top = ObjectShapeForm(VarObjectIndex).Top + TwipSize
        DisplayPosition ObjectShape(VarObjectIndex).Top, ObjectShape(VarObjectIndex).Left
        DisplayShape ObjectShape(VarObjectIndex).Top, ObjectShape(VarObjectIndex).Left, ObjectShape(VarObjectIndex).Width, ObjectShape(VarObjectIndex).Height
    End If

End Function

Private Sub mnuEditDown_Click()

    MoveObjectDown

End Sub

Function MoveObjectLeft()

    'move selected object to left
    If VarObjectType = "text" Then
        ObjectText(VarObjectIndex).Left = ObjectText(VarObjectIndex).Left - TwipSize
        DisplayPosition ObjectText(VarObjectIndex).Top, ObjectText(VarObjectIndex).Left
        DisplayShape ObjectText(VarObjectIndex).Top, ObjectText(VarObjectIndex).Left, ObjectText(VarObjectIndex).Width, ObjectText(VarObjectIndex).Height
    ElseIf VarObjectType = "picture" Then
        ObjectImage(VarObjectIndex).Left = ObjectImage(VarObjectIndex).Left - TwipSize
        DisplayPosition ObjectImage(VarObjectIndex).Top, ObjectImage(VarObjectIndex).Left
        DisplayShape ObjectImage(VarObjectIndex).Top, ObjectImage(VarObjectIndex).Left, ObjectImage(VarObjectIndex).Width, ObjectImage(VarObjectIndex).Height
    ElseIf VarObjectType = "shape" Then
        ObjectShape(VarObjectIndex).Left = ObjectShape(VarObjectIndex).Left - TwipSize
        ObjectShapeForm(VarObjectIndex).Left = ObjectShapeForm(VarObjectIndex).Left - TwipSize
        DisplayPosition ObjectShape(VarObjectIndex).Top, ObjectShape(VarObjectIndex).Left
        DisplayShape ObjectShape(VarObjectIndex).Top, ObjectShape(VarObjectIndex).Left, ObjectShape(VarObjectIndex).Width, ObjectShape(VarObjectIndex).Height
    End If

End Function

Private Sub mnuEditLeft_Click()

    MoveObjectLeft

End Sub

Function MoveObjectRight()

    'move selected object to right
    If VarObjectType = "text" Then
        ObjectText(VarObjectIndex).Left = ObjectText(VarObjectIndex).Left + TwipSize
        DisplayPosition ObjectText(VarObjectIndex).Top, ObjectText(VarObjectIndex).Left
        DisplayShape ObjectText(VarObjectIndex).Top, ObjectText(VarObjectIndex).Left, ObjectText(VarObjectIndex).Width, ObjectText(VarObjectIndex).Height
    ElseIf VarObjectType = "picture" Then
        ObjectImage(VarObjectIndex).Left = ObjectImage(VarObjectIndex).Left + TwipSize
        DisplayPosition ObjectImage(VarObjectIndex).Top, ObjectImage(VarObjectIndex).Left
        DisplayShape ObjectImage(VarObjectIndex).Top, ObjectImage(VarObjectIndex).Left, ObjectImage(VarObjectIndex).Width, ObjectImage(VarObjectIndex).Height
    ElseIf VarObjectType = "shape" Then
        ObjectShape(VarObjectIndex).Left = ObjectShape(VarObjectIndex).Left + TwipSize
        ObjectShapeForm(VarObjectIndex).Left = ObjectShapeForm(VarObjectIndex).Left + TwipSize
        DisplayPosition ObjectShape(VarObjectIndex).Top, ObjectShape(VarObjectIndex).Left
        DisplayShape ObjectShape(VarObjectIndex).Top, ObjectShape(VarObjectIndex).Left, ObjectShape(VarObjectIndex).Width, ObjectShape(VarObjectIndex).Height
    End If

End Function

Private Sub mnuEditRight_Click()

    MoveObjectRight

End Sub

Private Sub mnuEditUnselect_Click()

    ResetObjects

End Sub

Function MoveObjectUp()

    'move selected object up
    If VarObjectType = "text" Then
        ObjectText(VarObjectIndex).Top = ObjectText(VarObjectIndex).Top - TwipSize
        DisplayPosition ObjectText(VarObjectIndex).Top, ObjectText(VarObjectIndex).Left
        DisplayShape ObjectText(VarObjectIndex).Top, ObjectText(VarObjectIndex).Left, ObjectText(VarObjectIndex).Width, ObjectText(VarObjectIndex).Height
    ElseIf VarObjectType = "picture" Then
        ObjectImage(VarObjectIndex).Top = ObjectImage(VarObjectIndex).Top - TwipSize
        DisplayPosition ObjectImage(VarObjectIndex).Top, ObjectImage(VarObjectIndex).Left
        DisplayShape ObjectImage(VarObjectIndex).Top, ObjectImage(VarObjectIndex).Left, ObjectImage(VarObjectIndex).Width, ObjectImage(VarObjectIndex).Height
    ElseIf VarObjectType = "shape" Then
        ObjectShape(VarObjectIndex).Top = ObjectShape(VarObjectIndex).Top - TwipSize
        ObjectShapeForm(VarObjectIndex).Top = ObjectShapeForm(VarObjectIndex).Top - TwipSize
        DisplayPosition ObjectShape(VarObjectIndex).Top, ObjectShape(VarObjectIndex).Left
        DisplayShape ObjectShape(VarObjectIndex).Top, ObjectShape(VarObjectIndex).Left, ObjectShape(VarObjectIndex).Width, ObjectShape(VarObjectIndex).Height
    End If

End Function

Private Sub mnuEditUp_Click()

    MoveObjectUp

End Sub

Private Sub mnuFileExit_Click()

    End

End Sub

Function WriteFile(FileName As String)

    Me.MousePointer = vbHourglass
    Me.Enabled = False

    'write data to file
    Open FileName For Output As #1
        'write application properties
        Print #1, "Type=App|" & App.Major & "." & App.Minor & "." & App.Revision & "|" & Format(Now, "yyyy-mm-dd hh:nn:ss")
        'write master properties
        Print #1, "Type=Master|" & label.BackColor & "|" & label.BorderStyle & "|" & label.Width & "|" & label.Height
        'write picture objects
        For i = 1 To ObjectImage.UBound
            If ObjectImage(i).Visible = True Then 'only save of object is visible!
                Print #1, "Type=Picture|" & ObjectImageURL(i).Caption & "|" & ObjectImage(i).Left & "|" & ObjectImage(i).Top
            End If
        Next i
        'write text objects
        For i = 1 To ObjectText.UBound
            If ObjectText(i).Visible = True Then 'only save of object is visible!
                Print #1, "Type=Text|" & Replace(ObjectText(i).Caption, vbCrLf, "\nl\") & "|" & ObjectText(i).ForeColor & "|" & ObjectText(i).BackColor & "|" & ObjectText(i).BackStyle & "|" & ObjectText(i).FontName & "|" & ObjectText(i).FontSize & "|" & ObjectText(i).FontBold & "|" & ObjectText(i).FontItalic & "|" & ObjectText(i).FontUnderline & "|" & ObjectText(i).FontStrikethru & "|" & ObjectText(i).Left & "|" & ObjectText(i).Top
            End If
        Next i
        'write shape objects
        For i = 1 To ObjectShapeForm.UBound
            If ObjectShapeForm(i).Visible = True Then 'only save of object is visible!
                Print #1, "Type=Shape|" & ObjectShapeForm(i).FillColor & "|" & ObjectShapeForm(i).BorderColor & "|" & ObjectShapeForm(i).BorderStyle & "|" & ObjectShapeForm(i).Width & "|" & ObjectShapeForm(i).Height & "|" & ObjectShapeForm(i).Shape & "|" & ObjectShapeForm(i).Left & "|" & ObjectShapeForm(i).Top
            End If
        Next i
    Close #1
    
    Me.Enabled = True
    Me.MousePointer = vbNormal

End Function

Private Sub mnuFileExport_Click()

    DrawPicture
    DrawPicture
    Export.Show
    Me.Enabled = False

End Sub

Private Sub mnuFileNew_Click()

    If MsgBox("Are you sure you would like to create a new project?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        
        'unload objects
        For i = 1 To ObjectText.UBound
            Unload ObjectText(i)
            Unload ObjectTextList(i)
        Next i
        For i = 1 To ObjectImage.UBound
            Unload ObjectImage(i)
            Unload ObjectImageURL(i)
            Unload ObjectImageList(i)
        Next i
        
        'clear ListView
        ListView1.ListItems.Clear
        
        StatusBar1.Panels(1).Text = "Untitled"
        mnuFileSave.Caption = "&Save..."
        mnuFileSave.Enabled = False
        mnuFileSaveAs.Caption = "Save &As..."
        
    End If

End Sub

Function OpenFile(FileName As String, FileTitle As String)

        mnuFileSave.Enabled = True
        StatusBar1.Panels(1).Text = "Loading " & FileTitle & "..."
        
        label.Visible = False
        LabelShadow.Visible = False
        Me.Enabled = False
        Me.MousePointer = vbHourglass
        
        'unload objects
        For i = 1 To ObjectText.UBound
            Unload ObjectText(i)
            Unload ObjectTextList(i)
        Next i
        For i = 1 To ObjectImage.UBound
            Unload ObjectImage(i)
            Unload ObjectImageURL(i)
            Unload ObjectImageList(i)
        Next i
        For i = 1 To ObjectShapeForm.UBound
            Unload ObjectShape(i)
            Unload ObjectShapeForm(i)
            Unload ObjectShapeList(i)
        Next i
        
        'clear ListView1
        ListView1.ListItems.Clear
        
        'load data
        Open FileName For Input As #1
            Do While Not EOF(1)
                Line Input #1, FileContent
                'MASTER
                If Left(FileContent, 11) = "Type=Master" Then
                    ArrayLineContent = Split(FileContent, "|")
                    'background color
                    label.BackColor = ArrayLineContent(1)
                    'border
                    label.BorderStyle = ArrayLineContent(2)
                    'width and height
                    label.Width = ArrayLineContent(3)
                    label.Height = ArrayLineContent(4)
                    LabelShadow.Width = ArrayLineContent(3)
                    LabelShadow.Height = ArrayLineContent(4)
                'PICTURE OBJECT
                ElseIf Left(FileContent, 12) = "Type=Picture" Then
                    ArrayLineContent = Split(FileContent, "|")
                    SourceFile = ArrayLineContent(1)
                    If FileCheck(SourceFile) = True Then
                        Load ObjectImage(ObjectImage.Count)
                        Load ObjectImageURL(ObjectImageURL.Count)
                        Load ObjectImageList(ObjectImageList.Count)
                        ListView1.ListItems.Add , , "Picture " & ObjectImage.UBound, , 2
                        ListView1.ListItems(ListView1.ListItems.Count).Checked = True
                        'set properties
                        With ObjectImage(ObjectImage.UBound)
                            .Visible = True
                            .Picture = LoadPicture(ArrayLineContent(1))
                            .ZOrder (0)
                            .Left = ArrayLineContent(2)
                            .Top = ArrayLineContent(3)
                        End With
                        With ObjectImageURL(ObjectImageURL.UBound)
                            .Caption = ArrayLineContent(1)
                        End With
                        With ObjectImageList(ObjectImageList.UBound)
                            .Caption = ListView1.ListItems.Count
                        End With
                    Else
                        MsgBox "Picture could not be found:" & vbCrLf & SourceFile & vbCrLf & vbCrLf & "Picture Object was not created!", vbCritical + vbOKOnly, "Open File"
                    End If
                'SHAPE OBJECT
                ElseIf Left(FileContent, 10) = "Type=Shape" Then
                    ArrayLineContent = Split(FileContent, "|")
                    Load ObjectShape(ObjectShape.Count)
                    Load ObjectShapeForm(ObjectShapeForm.Count)
                    Load ObjectShapeList(ObjectShapeList.Count)
                    ListView1.ListItems.Add , , "Shape " & ObjectShape.UBound, , 3
                    ListView1.ListItems(ListView1.ListItems.Count).Checked = True
                    With ObjectShapeForm(ObjectShapeForm.UBound)
                        .Visible = True
                        .Width = ArrayLineContent(4)
                        .Height = ArrayLineContent(5)
                        .BorderColor = ArrayLineContent(2)
                        .FillColor = ArrayLineContent(1)
                        .BackStyle = ArrayLineContent(3)
                        .Shape = ArrayLineContent(6)
                        .Left = ArrayLineContent(7)
                        .Top = ArrayLineContent(8)
                        ZOrder (0)
                    End With
                    With ObjectShape(ObjectShape.UBound)
                        .Width = ArrayLineContent(4)
                        .Height = ArrayLineContent(5)
                        .Left = ArrayLineContent(7)
                        .Top = ArrayLineContent(8)
                        .Visible = True
                    End With
                    With ObjectShapeList(ObjectShapeList.UBound)
                        .Caption = ListView1.ListItems.Count
                    End With
                'TEXT OBJECT
                ElseIf Left(FileContent, 9) = "Type=Text" Then
                    ArrayLineContent = Split(FileContent, "|")
                    Load ObjectText(ObjectText.Count)
                    Load ObjectTextList(ObjectTextList.Count)
                    ListView1.ListItems.Add , , "Text " & ObjectText.UBound, , 1
                    ListView1.ListItems(ListView1.ListItems.Count).Checked = True
                    'set properties
                    With ObjectText(ObjectText.UBound)
                        .Visible = True
                        .Caption = Replace(ArrayLineContent(1), "\nl\", vbCrLf)
                        .FontName = ArrayLineContent(5)
                        .FontSize = ArrayLineContent(6)
                        .ForeColor = ArrayLineContent(2)
                        .FontBold = ArrayLineContent(7)
                        .FontItalic = ArrayLineContent(8)
                        .FontUnderline = ArrayLineContent(9)
                        .FontStrikethru = ArrayLineContent(10)
                        .BackStyle = ArrayLineContent(4)
                        .BackColor = ArrayLineContent(3)
                        .ZOrder (0)
                        .Left = ArrayLineContent(11)
                        .Top = ArrayLineContent(12)
                    End With
                    With ObjectTextList(ObjectTextList.UBound)
                        .Caption = ListView1.ListItems.Count
                    End With
                End If
            Loop
        Close #1
        
        ResizeElements
        
        StatusBar1.Panels(1).Text = FileName
        mnuFileSave.Caption = "&Save " & FileTitle & "..."
        mnuFileSave.Enabled = True
        mnuFileSaveAs.Caption = "&Save " & FileTitle & " As..."

        label.Visible = True
        LabelShadow.Visible = False
        Me.Enabled = True
        Me.MousePointer = vbNormal

End Function

Private Sub mnuFileOpen_Click()

    On Error GoTo Err:

    With CommonDialog1
        .Filter = "FastLabel File (*.fal)|*.fal|"
        .CancelError = True
        .ShowOpen
        OpenFile .FileName, .FileTitle
    End With

Err:
    Exit Sub

End Sub

Private Sub mnuFilePrint_Click()

    DrawPicture
    DrawPicture
    Me.Enabled = False
    PrintManager.Show

End Sub

Private Sub mnuFileSave_Click()

    WriteFile (StatusBar1.Panels(1).Text)

End Sub

Private Sub mnuFileSaveAs_Click()

    On Error GoTo Err:

    With CommonDialog1
        .Filter = "FastLabel File (*.fal)|*.fal|"
        .CancelError = True
        .Flags = cdlOFNOverwritePrompt
        .ShowSave
        mnuFileSave.Enabled = True
        StatusBar1.Panels(1).Text = .FileName
        mnuFileSave.Caption = "&Save " & .FileTitle & "..."
        mnuFileSaveAs.Caption = "&Save " & .FileTitle & " As..."
        mnuFileSave.Enabled = True
        WriteFile (.FileName)
    End With

Err:
    Exit Sub

End Sub

Private Sub mnuHelpAbout_Click()

    About.Show
    Me.Enabled = False

End Sub

Private Sub mnuInserShape_Click()

    Me.Enabled = False
    InsertShape.Show

End Sub

Private Sub mnuInsertPicture_Click()

    Me.Enabled = False
    InsertPicture.Show

End Sub

Private Sub mnuInsertProperties_Click()

    Me.Enabled = False
    Form2.Show

End Sub

Private Sub mnuInsertText_Click()

    Me.Enabled = False
    InsertText.Show

End Sub

Function DisplayShape(Top As String, Left As String, Width As String, Height As String)

    DisplayObjectShape.Top = Top - 0.05
    DisplayObjectShape.Left = Left - 0.1
    DisplayObjectShape.Width = Width + 0.2
    DisplayObjectShape.Height = Height + 0.1
    DisplayObjectShape.Visible = True
    DisplayObjectShape.ZOrder (0)

End Function

Function DisplayPosition(Top As String, Left As String)

    If mnuViewSizeLabel.Checked = True Then
        DisplayMovePosition.Top = Top
        DisplayMovePosition.Left = Left
        DisplayMovePosition.Caption = Round(DisplayMovePosition.Left, 2) & " x " & Round(DisplayMovePosition.Top, 2)
        DisplayMovePosition.Visible = True
        DisplayMovePosition.ZOrder (0)
    End If

End Function

Function SelectObject()

    'enable move menu
    mnuEditMove.Enabled = True
    mnuEditBringtofront.Enabled = True
    mnuEditSendtoback.Enabled = True
    mnuEditUnselect.Enabled = True
    mnuEditDelete.Enabled = True
    mnuEditModify.Enabled = True

    If VarObjectType = "text" Then
        StatusBar1.Panels(4).Text = "Object: " & Round(ObjectText(VarObjectIndex).Width, 2) & " x " & Round(ObjectText(VarObjectIndex).Height, 2)
        DisplayPosition ObjectText(VarObjectIndex).Top, ObjectText(VarObjectIndex).Left
        DisplayShape ObjectText(VarObjectIndex).Top, ObjectText(VarObjectIndex).Left, ObjectText(VarObjectIndex).Width, ObjectText(VarObjectIndex).Height
    ElseIf VarObjectType = "picture" Then
        StatusBar1.Panels(4).Text = "Object: " & Round(ObjectImage(VarObjectIndex).Width, 2) & " x " & Round(ObjectImage(VarObjectIndex).Height, 2)
        DisplayPosition ObjectImage(VarObjectIndex).Top, ObjectImage(VarObjectIndex).Left
        DisplayShape ObjectImage(VarObjectIndex).Top, ObjectImage(VarObjectIndex).Left, ObjectImage(VarObjectIndex).Width, ObjectImage(VarObjectIndex).Height
    ElseIf VarObjectType = "shape" Then
        StatusBar1.Panels(4).Text = "Object: " & Round(ObjectShape(VarObjectIndex).Width, 2) & " x " & Round(ObjectShape(VarObjectIndex).Height, 2)
        DisplayPosition ObjectShape(VarObjectIndex).Top, ObjectShape(VarObjectIndex).Left
        DisplayShape ObjectShape(VarObjectIndex).Top, ObjectShape(VarObjectIndex).Left, ObjectShape(VarObjectIndex).Width, ObjectShape(VarObjectIndex).Height
    End If
    
    ObjectIsSelected = True

End Function

Private Sub mnuViewCenterLines_Click()

    If mnuViewCenterLines.Checked = True Then
        mnuViewCenterLines.Checked = False
        HorizontalCenterLine.Visible = False
        VerticalCenterLine.Visible = False
    Else
        mnuViewCenterLines.Checked = True
        HorizontalCenterLine.Visible = True
        VerticalCenterLine.Visible = True
    End If

End Sub

Private Sub mnuViewMouseTracker_Click()

    If mnuViewMouseTracker.Checked = True Then
        mnuViewMouseTracker.Checked = False
    Else
        mnuViewMouseTracker.Checked = True
    End If

End Sub

Private Sub mnuViewSizeLabel_Click()

    If mnuViewSizeLabel.Checked = True Then
        mnuViewSizeLabel.Checked = False
    Else
        mnuViewSizeLabel.Checked = True
    End If

End Sub

Private Sub ObjectImage_Click(Index As Integer)

    On Error Resume Next
    
    ResetObjects

    'set properties
    VarObjectType = "picture"
    VarObjectIndex = Index
    VarListIndex = ObjectImageList(Index).Caption
    
    ListView1.ListItems.Item(VarListIndex).Bold = True
    ListView1.ListItems.Item(VarListIndex).Selected = True

    SelectObject

End Sub

Private Sub ObjectImage_DblClick(Index As Integer)

    InsertPicture.Show
    InsertPicture.HiddenObjectIndex.Caption = Index
    InsertPicture.Text1.Text = ObjectImageURL(Index).Caption
    InsertPicture.Image1.Picture = ObjectImage(Index).Picture
    Me.Enabled = False

End Sub

Private Sub ObjectImage_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    MoveObject = True

End Sub

Private Sub ObjectImage_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    MoveObject = False
    If ObjectIsSelected = True Then
        If Button = 2 Then
            PopupMenu mnuEdit
        End If
    End If

End Sub

Private Sub ObjectShape_Click(Index As Integer)

    On Error Resume Next

    ResetObjects
    
    'set properties
    VarObjectType = "shape"
    VarObjectIndex = Index
    VarListIndex = ObjectShapeList(Index).Caption
    
    ListView1.ListItems.Item(VarListIndex).Bold = True
    ListView1.ListItems.Item(VarListIndex).Selected = True
    
    SelectObject

End Sub

Private Sub ObjectShape_DblClick(Index As Integer)


    InsertShape.Show
    InsertShape.HiddenObjectIndex.Caption = Index
    
    If ObjectShapeForm(Index).Shape = 0 Then
        InsertShape.Combo1.Text = "Rectangle"
    ElseIf ObjectShapeForm(Index).Shape = 3 Then
        InsertShape.Combo1.Text = "Circle"
    End If
    
    InsertShape.Text1.Text = ObjectShapeForm(Index).Width
    InsertShape.Text2.Text = ObjectShapeForm(Index).Height
    InsertShape.HScroll1.Value = Round(ObjectShapeForm(Index).Width * 100)
    InsertShape.HScroll2.Value = Round(ObjectShapeForm(Index).Height * 100)
    If ObjectShapeForm(Index).BorderStyle = 1 Then
        InsertShape.Check1.Value = 1
    Else
        InsertShape.Check1.Value = 0
    End If
    
    InsertShape.Picture1.BackColor = ObjectShapeForm(Index).FillColor
    InsertShape.Picture2.BackColor = ObjectShapeForm(Index).BorderColor

End Sub

Private Sub ObjectShape_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    MoveObject = True
    SetCurrentX = HiddenX
    SetCurrentY = HiddenY

End Sub

Private Sub ObjectShape_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    MoveObject = False
    If ObjectIsSelected = True Then
        If Button = 2 Then
            PopupMenu mnuEdit
        End If
    End If

End Sub

Private Sub ObjectText_Click(Index As Integer)

    On Error Resume Next

    ResetObjects
    
    'set properties
    VarObjectType = "text"
    VarObjectIndex = Index
    VarListIndex = ObjectTextList(Index).Caption
    
    ListView1.ListItems.Item(VarListIndex).Bold = True
    ListView1.ListItems.Item(VarListIndex).Selected = True
    
    SelectObject

End Sub

Private Sub ObjectText_DblClick(Index As Integer)

    InsertText.Show
    InsertText.HiddenObjectIndex.Caption = Index
    InsertText.Text1.Text = ObjectText(Index).Caption
    InsertText.Combo1.Text = ObjectText(Index).FontName
    InsertText.Combo2.Text = ObjectText(Index).FontSize
    
    If ObjectText(Index).FontBold = True Then
        InsertText.Check1.Value = 1
    End If
    If ObjectText(Index).FontItalic = True Then
        InsertText.Check2.Value = 1
    End If
    If ObjectText(Index).FontUnderline = True Then
        InsertText.Check3.Value = 1
    End If
    If ObjectText(Index).FontStrikethru = True Then
        InsertText.Check4.Value = 1
    End If
    
    InsertText.Picture1.BackColor = ObjectText(Index).ForeColor

End Sub

Private Sub ObjectText_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    MoveObject = True
    SetCurrentX = HiddenX
    SetCurrentY = HiddenY

End Sub

Private Sub ObjectText_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If MoveObject = True Then
        
    End If

End Sub

Private Sub ObjectText_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    MoveObject = False
    If ObjectIsSelected = True Then
        If Button = 2 Then
            PopupMenu mnuEdit
        End If
    End If
   
End Sub

Private Sub WorkBar_Click()
    ResetObjects
End Sub

Private Sub WorkBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    StatusBar1.Panels(3).Text = "Position: 0 x 0"
    HorizontalSplit.Visible = False
    VerticalSplit.Visible = False
    BarTopLine.Visible = False
    BarLeftLine.Visible = False

End Sub

Private Sub WorkBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        PopupMenu mnuInsert
    End If

End Sub
