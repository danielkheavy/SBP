VERSION 5.00
Begin VB.Form frmFonts 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fonts"
   ClientHeight    =   4920
   ClientLeft      =   2250
   ClientTop       =   2310
   ClientWidth     =   5745
   Icon            =   "frmFonts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picFrame 
      BorderStyle     =   0  'None
      Height          =   4545
      Left            =   165
      ScaleHeight     =   4545
      ScaleWidth      =   5475
      TabIndex        =   0
      Top             =   60
      Width           =   5475
      Begin PrintLabels.chameleonButton cmdOk 
         Height          =   495
         Left            =   4125
         TabIndex        =   16
         Top             =   3975
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   873
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFonts.frx":0E52
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ListBox List1 
         Height          =   2205
         Left            =   15
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   600
         Width           =   3060
      End
      Begin VB.ComboBox cboFind 
         Height          =   315
         Left            =   15
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   8
         Text            =   "Combo2"
         Top             =   240
         Width           =   3060
      End
      Begin VB.ListBox ListSize 
         Height          =   2205
         Left            =   3225
         TabIndex        =   7
         Top             =   615
         Width           =   735
      End
      Begin VB.OptionButton optStyle 
         Caption         =   "Bold Italic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "bolditalic"
         Top             =   1605
         Width           =   1320
      End
      Begin VB.OptionButton optStyle 
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
         Height          =   345
         Index           =   2
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "bold"
         Top             =   1260
         Width           =   1320
      End
      Begin VB.OptionButton optStyle 
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
         Height          =   345
         Index           =   1
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "italic"
         Top             =   915
         Width           =   1320
      End
      Begin VB.OptionButton optStyle 
         Caption         =   "Regular"
         Height          =   345
         Index           =   0
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "regular"
         Top             =   570
         Value           =   -1  'True
         Width           =   1320
      End
      Begin VB.CheckBox chkStrikethru 
         Caption         =   "Strikethru"
         Height          =   345
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2415
         Width           =   1320
      End
      Begin VB.CheckBox chkUnderline 
         Caption         =   "Underline"
         Height          =   345
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2070
         Width           =   1320
      End
      Begin VB.Label lblFontName 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblFontName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   15
         TabIndex        =   15
         Top             =   2925
         Width           =   5385
      End
      Begin VB.Label lblFontSize 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3225
         TabIndex        =   14
         Top             =   255
         Width           =   735
      End
      Begin VB.Label lblSty 
         BackStyle       =   0  'Transparent
         Caption         =   "Style:"
         Height          =   210
         Left            =   4065
         TabIndex        =   13
         Top             =   15
         Width           =   495
      End
      Begin VB.Label lblStyle 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Regular"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4080
         TabIndex        =   12
         Top             =   285
         Width           =   1320
      End
      Begin VB.Label lblSize 
         BackStyle       =   0  'Transparent
         Caption         =   "Size:"
         Height          =   255
         Left            =   3360
         TabIndex        =   11
         Top             =   0
         Width           =   495
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Font:"
         Height          =   195
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmFonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Font_Style            As String
Public Font_Size             As Long
Public Font_Name             As String
Public Font_Underline        As Boolean
Public Font_StrikeThru       As Boolean
Public Font_Bold             As Boolean
Public Font_Italic           As Boolean

Private Sub cboFind_Click()

   If cboFind.ListCount > 0 Then List1.ListIndex = cboFind.ListIndex

End Sub

Private Sub cboFind_KeyPress(KeyAscii As Integer)

   cValidate.AutoMatch cboFind, KeyAscii

End Sub

Private Sub chkStrikethru_Click()

   On Error GoTo Err_Proc

   If chkStrikethru.Value Then
      Font_StrikeThru = True
   Else
      Font_StrikeThru = False
   End If
   Call UpdateSample

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmFonts", "chkStrikethru_Click"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub chkUnderline_Click()

   On Error GoTo Err_Proc

   If chkUnderline.Value Then
      Font_Underline = True
   Else
      Font_Underline = False
   End If
   Call UpdateSample

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmFonts", "chkUnderline_Click"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub cmdOK_Click()

   Me.Hide

End Sub

Private Sub Form_Load()

  Dim hdc        As Long
  Dim i          As Long
  Dim IndexFound As Boolean

   On Error GoTo Err_Proc


   Me.Move frmMain.Left, frmMain.Top, frmMain.Width, frmMain.Height
   cScreen.CenterObject Me, picFrame

   '/* Load Font Sizes */
   For i = 6 To 120
      ListSize.AddItem CStr(i)
   Next i

   '/* Load Font Names */
   hdc = GetDC(List1.hwnd)
   ShowFontType = 4 'True Type
   EnumFontFamilies hdc, vbNullString, AddressOf EnumFontFamTypeProc, List1
   ShowFontType = 1
   EnumFontFamilies hdc, vbNullString, AddressOf EnumFontFamTypeProc, List1
   ReleaseDC List1.hwnd, hdc

   hdc = GetDC(cboFind.hwnd)
   ShowFontType = 4 'True Type
   EnumFontFamilies hdc, vbNullString, AddressOf EnumFontFamTypeProc, cboFind
   ShowFontType = 1
   EnumFontFamilies hdc, vbNullString, AddressOf EnumFontFamTypeProc, cboFind
   ReleaseDC cboFind.hwnd, hdc

   '/* Set default Font name */
   If Font_Name > vbNullString Then
      For i = 0 To cboFind.ListCount - 1
         If cboFind.List(i) = Font_Name Then
            'cboFind.ListIndex = i
            List1.ListIndex = i
            IndexFound = True
            Exit For
         End If
      Next i
   Else
      'cboFind.ListIndex = 0
      List1.ListIndex = 0
   End If
   If Not IndexFound Then
      'cboFind.ListIndex = 0
      List1.ListIndex = 0
   End If

   '/* Set default size */
   If Font_Size > 0 Then
      ListSize.ListIndex = Font_Size - 6
   Else
      ListSize.ListIndex = 4
   End If

   '/* Set default style */
   If Font_Style > vbNullString Then
      Select Case Font_Style
      Case "regular"
         optStyle(0).Value = True
      Case "italic"
         optStyle(1).Value = True
      Case "bold"
         optStyle(2).Value = True
      Case "bolditalic"
         optStyle(3).Value = True
      End Select
   Else
      If Not Font_Bold And Not Font_Italic Then
         optStyle(0).Value = True
      ElseIf Not Font_Bold And Font_Italic Then
         optStyle(1).Value = True
      ElseIf Font_Bold And Not Font_Italic Then
         optStyle(2).Value = True
      ElseIf Font_Bold And Font_Italic Then
         optStyle(3).Value = True
      End If
   End If

   chkUnderline.Value = Font_Underline
   chkStrikethru.Value = Font_StrikeThru

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmFonts", "Form_Load"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   If UnloadMode = vbFormControlMenu Then
      Cancel = True
      Exit Sub
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

   On Error GoTo Err_Proc

   Font_Style = vbNullString
   Font_Size = 0
   Font_Name = vbNullString
   Font_Underline = False
   Font_StrikeThru = False
   Set frmFonts = Nothing

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmFonts", "Form_Unload"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub List1_Click()

   Font_Name = List1.Text

   cboFind.ListIndex = List1.ListIndex
   Call UpdateSample

End Sub

Private Sub ListSize_Click()

   Font_Size = CInt(ListSize.Text)

   lblFontSize = ListSize.Text
   Call UpdateSample

End Sub

Private Sub optStyle_Click(Index As Integer)

   Font_Style = optStyle(Index).Tag

   lblStyle = optStyle(Index).Caption
   Call UpdateSample

End Sub

Public Sub UpdateSample()

   On Local Error Resume Next

   With lblFontName
      .Caption = Font_Name
      .FontName = Font_Name
      .FontSize = Font_Size + 10
      .FontStrikethru = Font_StrikeThru
      .FontUnderline = Font_Underline
      Select Case Font_Style
      Case "regular"
         Font_Bold = False
         Font_Italic = False
         .FontBold = False
         .FontItalic = False
      Case "italic"
         Font_Bold = False
         Font_Italic = True
         .FontBold = False
         .FontItalic = True
      Case "bold"
         Font_Bold = True
         Font_Italic = False
         .FontBold = True
         .FontItalic = False
      Case "bolditalic"
         Font_Bold = True
         Font_Italic = True
         .FontBold = True
         .FontItalic = True
      End Select
      .Refresh
   End With
   cboFind.SetFocus
   On Local Error GoTo 0

End Sub

