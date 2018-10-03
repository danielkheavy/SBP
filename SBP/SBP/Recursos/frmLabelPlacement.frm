VERSION 5.00
Begin VB.Form frmLabelPlacement 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Label Placement"
   ClientHeight    =   5145
   ClientLeft      =   3240
   ClientTop       =   2325
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   Begin PrintLabels.chameleonButton cmdSaveAs 
      Height          =   390
      Left            =   3885
      TabIndex        =   17
      Top             =   3780
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   688
      BTYPE           =   3
      TX              =   "Save As"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLabelPlacement.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrintLabels.chameleonButton cmdFonts 
      Height          =   465
      Left            =   4770
      TabIndex        =   16
      Top             =   3015
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "Font Settings"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLabelPlacement.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txt_Down 
      Height          =   300
      Left            =   4035
      TabIndex        =   6
      Text            =   "10"
      Top             =   3300
      Width           =   465
   End
   Begin VB.TextBox txt_Across 
      Height          =   300
      Left            =   4035
      TabIndex        =   5
      Text            =   "3"
      Top             =   2925
      Width           =   465
   End
   Begin VB.TextBox txt_HorzPitch 
      Height          =   300
      Left            =   2520
      TabIndex        =   4
      Text            =   "2.5"
      Top             =   3300
      Width           =   465
   End
   Begin VB.TextBox txt_VerPitch 
      Height          =   300
      Left            =   2520
      TabIndex        =   3
      Text            =   "1"
      Top             =   2925
      Width           =   465
   End
   Begin VB.TextBox txt_SideMargin 
      Height          =   300
      Left            =   1065
      TabIndex        =   2
      Text            =   ".19"
      Top             =   3300
      Width           =   465
   End
   Begin VB.TextBox txt_TopMargin 
      Height          =   300
      Left            =   1065
      TabIndex        =   1
      Text            =   ".5"
      Top             =   2925
      Width           =   465
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1050
      Style           =   2  'Dropdown List
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3825
      Width           =   2745
   End
   Begin PrintLabels.chameleonButton cmdDelete 
      Height          =   390
      Left            =   4815
      TabIndex        =   18
      Top             =   3795
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   688
      BTYPE           =   3
      TX              =   "Delete"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLabelPlacement.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrintLabels.Frame3D Frame3D1 
      Height          =   2565
      Left            =   585
      Top             =   210
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   4524
      BorderType      =   5
      BevelWidth      =   3
      BevelInner      =   0
      Caption3D       =   0
      CaptionAlliment =   0
      CaptionLocation =   0
      BackColor       =   -2147483633
      CornerDiameter  =   7
      FillColor       =   16761247
      FillStyle       =   1
      FloodPercent    =   0
      FloodShowPct    =   0   'False
      FloodType       =   0
      FloodColor      =   16761247
      MousePointer    =   0
      MouseIcon       =   "frmLabelPlacement.frx":0054
      Picture         =   "frmLabelPlacement.frx":0070
      Border3DHighlight=   -2147483628
      Border3DShadow  =   -2147483632
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   -2147483630
      Caption         =   ""
      UseMnemonic     =   0   'False
   End
   Begin PrintLabels.Frame3D Picture2 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      Top             =   4485
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   1164
      BorderType      =   1
      BevelWidth      =   3
      BevelInner      =   0
      Caption3D       =   0
      CaptionAlliment =   0
      CaptionLocation =   0
      BackColor       =   -2147483633
      CornerDiameter  =   7
      FillColor       =   16761247
      FillStyle       =   1
      FloodPercent    =   0
      FloodShowPct    =   0   'False
      FloodType       =   0
      FloodColor      =   16761247
      MousePointer    =   0
      MouseIcon       =   "frmLabelPlacement.frx":2C3E6
      Picture         =   "frmLabelPlacement.frx":2C402
      Border3DHighlight=   -2147483628
      Border3DShadow  =   -2147483632
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   -2147483630
      Caption         =   ""
      UseMnemonic     =   0   'False
      Begin PrintLabels.chameleonButton cmd_LabelSizeOK 
         Height          =   555
         Left            =   4455
         TabIndex        =   19
         Top             =   60
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   "Ok"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   255
         MPTR            =   1
         MICON           =   "frmLabelPlacement.frx":2C41E
         PICN            =   "frmLabelPlacement.frx":2C43A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PrintLabels.chameleonButton cmd_LblSizeQuit 
         Height          =   555
         Left            =   5130
         TabIndex        =   20
         Top             =   60
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   "Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   65280
         MPTR            =   1
         MICON           =   "frmLabelPlacement.frx":2C900
         PICN            =   "frmLabelPlacement.frx":2C91C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " Font: "
      Height          =   195
      Index           =   6
      Left            =   555
      TabIndex        =   15
      Top             =   4200
      Width           =   450
   End
   Begin VB.Label lblFont 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1050
      TabIndex        =   14
      Top             =   4200
      Width           =   45
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "No.Down:"
      Height          =   195
      Index           =   5
      Left            =   3285
      TabIndex        =   13
      Top             =   3315
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "No.Across:"
      Height          =   195
      Index           =   4
      Left            =   3225
      TabIndex        =   12
      Top             =   2955
      Width           =   780
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Horz.Pitch:"
      Height          =   195
      Index           =   3
      Left            =   1710
      TabIndex        =   11
      Top             =   3330
      Width           =   780
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Ver.Pitch:"
      Height          =   195
      Index           =   2
      Left            =   1800
      TabIndex        =   10
      Top             =   2955
      Width           =   690
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Side Margin:"
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   9
      Top             =   3330
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Top Margin:"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   8
      Top             =   2955
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " Scheme: "
      Height          =   195
      Index           =   7
      Left            =   285
      TabIndex        =   7
      Top             =   3855
      Width           =   720
   End
End
Attribute VB_Name = "frmLabelPlacement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tFontSize             As Long
Private tFontName             As String
Private tFontStyle            As String
Private tFontUnderline        As Boolean
Private tFontStrikeThru       As Boolean

Private Sub cmd_LabelSizeOK_Click()

   On Error GoTo Err_Proc

   TopMargin = Val(txt_TopMargin)
   SideMargin = Val(txt_SideMargin)
   VerPitch = Val(txt_VerPitch)
   HorzPitch = Val(txt_HorzPitch)
   NoAcross = Val(txt_Across)
   NoDown = Val(txt_Down)

   dFontSize = tFontSize
   dFontName = tFontName
   dFontStyle = tFontStyle
   dFontUnderline = tFontUnderline
   dFontStrikeThru = tFontStrikeThru

   SchemeID = Combo1.ItemData(Combo1.ListIndex)
   dPrintScheme = Combo1.Text
   frmMain.lblPrintScheme = dPrintScheme
   frmMain.lblFont = dFontName & ", " & CStr(dFontSize)
   frmMain.lblFont.Font = dFontName

   SaveSetting App.Title, "User", "PrintScheme", dPrintScheme
   Unload Me

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmLabelPlace", "cmd_LabelSizeOK_Click"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub cmd_LblSizeQuit_Click()

   Unload Me

End Sub

Private Sub cmdDelete_Click()

  Dim Mydb  As ADODB.Connection
  Dim MySet As ADODB.Recordset

   On Error GoTo Err_Proc


   If Combo1.Text = "Default" Then
      MsgBox "You can not Delete this scheme", vbInformation
      Exit Sub
   End If

   Call OpenDB(Mydb, , dbPath & App.Title & ".dat")
   Call OpenRS(MySet, "Select * from Settings", Mydb)

   If ADOFindFirst(MySet, "[ID]=" & Format$(Combo1.ItemData(Combo1.ListIndex))) Then
      MySet.Delete
   End If
   MySet.Close
   Mydb.Close
   Call FillCombo

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmLabelPlace", "cmdDelete_Click"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub cmdFonts_Click()

   On Error GoTo Err_Proc

   With frmFonts
      .Font_Style = tFontStyle
      .Font_Size = tFontSize
      .Font_Name = tFontName
      .Font_Underline = tFontUnderline
      .Font_StrikeThru = tFontStrikeThru

      .Show vbModal

      If Not QuitCommand Then
         tFontStyle = .Font_Style
         tFontSize = .Font_Size
         tFontName = .Font_Name
         tFontUnderline = .Font_Underline
         tFontStrikeThru = .Font_StrikeThru
         lblFont = tFontName & ", " & CStr(tFontSize)
         lblFont.Font = tFontName
      End If
   End With
   Unload frmFonts
   QuitCommand = False

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmLabelPlace", "cmdFonts_Click"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub cmdSaveAs_Click()

   Dim Result    As String
   Dim AddingNew As Boolean
   Dim Mydb      As ADODB.Connection
   Dim MySet     As ADODB.Recordset
   Dim i         As Long
   Dim n         As Long

   On Error GoTo Err_Proc


   Call OpenDB(Mydb, , AppPath & App.Title & ".dat")
   Call OpenRS(MySet, "Select * from Settings", Mydb)

   Result = InputBox("Save Label Scheme as: ", "Save settings", Combo1)
   Result = Trim$(Result)

   If Result > vbNullString Then
      cValidate.FcaseString Result
      If Not ADOFindFirst(MySet, "[Description]='" & Result & "'") Then
         MySet.AddNew
         AddingNew = True
      Else
         SchemeID = MySet!ID
      End If
      MySet!Description = Result
      MySet!TopMargin = Val(txt_TopMargin)
      MySet!SideMargin = Val(txt_SideMargin)
      MySet!VPitch = Val(txt_VerPitch)
      MySet!HPitch = Val(txt_HorzPitch)
      MySet!NoAcross = Val(txt_Across)
      MySet!NoDown = Val(txt_Down)
      MySet!FontSize = tFontSize
      MySet!FontName = tFontName
      MySet!FontStyle = tFontStyle
      MySet!FontUnderline = tFontUnderline
      MySet!FontStrikethru = tFontStrikeThru
      MySet.Update
      dPrintScheme = Result
      SaveSetting App.Title, "User", "PrintScheme", dPrintScheme
   End If
   If AddingNew Then
      MySet.MoveLast
      SchemeID = MySet!ID
   End If

   MySet.Close
   Mydb.Close
   AddingNew = False

   Call FillCombo

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmLabelPlace", "cmdSaveAs_Click"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub Combo1_Click()

   Call UpdateFields

End Sub

Private Sub FillCombo()

   Dim Mydb  As ADODB.Connection
   Dim MySet As ADODB.Recordset
   Dim i     As Long

   On Error GoTo Err_Proc


   Call OpenDB(Mydb, , AppPath & App.Title & ".dat")
   Call OpenRS(MySet, "Select * from Settings", Mydb)

   Combo1.Clear
   If Not (MySet.EOF And MySet.BOF) Then
      Do
         Combo1.AddItem MySet!Description
         Combo1.ItemData(Combo1.NewIndex) = MySet!ID
         MySet.MoveNext
      Loop Until MySet.EOF

      Combo1.ListIndex = False
      For i = 0 To Combo1.ListCount - 1
         If Combo1.ItemData(i) = SchemeID Then
            Combo1.ListIndex = i
            Exit For
         End If
      Next i

   End If
   MySet.Close
   Mydb.Close

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmLabelPlace", "FillCombo"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub Form_Load()

   On Error GoTo Err_Proc

   Me.Move frmMain.Left, frmMain.Top, frmMain.Width, frmMain.Height
   Me.Icon = frmMain.Icon
   QuitCommand = False

   txt_TopMargin.Text = Format$(TopMargin)
   txt_SideMargin.Text = Format$(SideMargin)
   txt_VerPitch.Text = Format$(VerPitch)
   txt_HorzPitch.Text = Format$(HorzPitch)
   txt_Across.Text = Format$(NoAcross)
   txt_Down.Text = Format$(NoDown)
   tFontSize = dFontSize
   tFontName = dFontName
   tFontStyle = dFontStyle
   tFontUnderline = dFontUnderline
   tFontStrikeThru = dFontStrikeThru

   Call FillCombo

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmLabelPlace", "Form_Load"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set frmLabelPlacement = Nothing

End Sub

Private Sub txt_Across_GotFocus()

   txtGotFocus txt_Across

End Sub

Private Sub txt_Down_GotFocus()

   txtGotFocus txt_Down

End Sub

Private Sub txt_HorzPitch_GotFocus()

   txtGotFocus txt_HorzPitch

End Sub

Private Sub txt_SideMargin_GotFocus()

   txtGotFocus txt_SideMargin

End Sub

Private Sub txt_TopMargin_GotFocus()

   txtGotFocus txt_TopMargin

End Sub

Private Sub txt_VerPitch_GotFocus()

   txtGotFocus txt_VerPitch

End Sub

Private Sub UpdateFields()

  Dim Mydb  As ADODB.Connection
  Dim MySet As ADODB.Recordset

   On Error GoTo Err_Proc


   Call OpenDB(Mydb, , AppPath & App.Title & ".dat")
   Call OpenRS(MySet, "Select * from Settings", Mydb)

   If Not (MySet.EOF And MySet.BOF) Then
      ADOFindFirst MySet, "[ID]=" & Format$(Combo1.ItemData(Combo1.ListIndex))
      txt_TopMargin = MySet!TopMargin
      txt_SideMargin = MySet!SideMargin
      txt_VerPitch = MySet!VPitch
      txt_HorzPitch = MySet!HPitch
      txt_Across = MySet!NoAcross
      txt_Down = MySet!NoDown
      tFontSize = MySet!FontSize
      tFontName = MySet!FontName
      tFontStyle = MySet!FontStyle
      tFontUnderline = MySet!FontUnderline
      tFontStrikeThru = MySet!FontStrikethru
   End If
   MySet.Close
   Mydb.Close

   lblFont = tFontName & ", " & CStr(tFontSize)
   lblFont.Font = tFontName

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmLabelPlace", "UpdateFields"
   Err.Clear
   Resume Exit_Here

End Sub

