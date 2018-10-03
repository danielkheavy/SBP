VERSION 5.00
Begin VB.Form frmDelimiters 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2535
   ClientLeft      =   6390
   ClientTop       =   3510
   ClientWidth     =   2985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmDelimiters.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   Begin PrintLabels.chameleonButton cmdOk 
      Height          =   405
      Left            =   1965
      TabIndex        =   0
      Top             =   1815
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   714
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
      MICON           =   "frmDelimiters.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrintLabels.chameleonButton cmdQuit 
      Height          =   405
      Left            =   1275
      TabIndex        =   1
      Top             =   1815
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   714
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmDelimiters.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrintLabels.Frame3D Frame1 
      Height          =   1440
      Left            =   360
      Top             =   210
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   2540
      BorderType      =   1
      BevelWidth      =   3
      BevelInner      =   0
      Caption3D       =   1
      CaptionAlliment =   0
      CaptionLocation =   1
      BackColor       =   -2147483633
      CornerDiameter  =   7
      FillColor       =   16761247
      FillStyle       =   1
      FloodPercent    =   0
      FloodShowPct    =   0   'False
      FloodType       =   0
      FloodColor      =   16761247
      MousePointer    =   0
      MouseIcon       =   "frmDelimiters.frx":0044
      Picture         =   "frmDelimiters.frx":0060
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
      Caption         =   "Delimiters"
      UseMnemonic     =   0   'False
      Begin VB.OptionButton optCrLf 
         Caption         =   "Carriage Return"
         Height          =   195
         Left            =   285
         TabIndex        =   5
         Top             =   1110
         Value           =   -1  'True
         Width           =   1530
      End
      Begin VB.OptionButton optComma 
         Caption         =   "Comma"
         Height          =   195
         Left            =   285
         TabIndex        =   4
         Top             =   840
         Width           =   1530
      End
      Begin VB.OptionButton optSemicolon 
         Caption         =   "Semicolon"
         Height          =   195
         Left            =   285
         TabIndex        =   3
         Top             =   570
         Width           =   1530
      End
      Begin VB.OptionButton optTab 
         Caption         =   "Tab"
         Height          =   195
         Left            =   285
         TabIndex        =   2
         Top             =   300
         Width           =   1530
      End
   End
End
Attribute VB_Name = "frmDelimiters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ImportFileName As String

Private Sub cmdOK_Click()

   Me.Hide

End Sub

Private Sub cmdQuit_Click()

   QuitCommand = True

   Me.Hide

End Sub

Private Sub Form_Load()

   Dim tString As String
   Dim tMatch  As Boolean
   Dim FF      As Long

   On Error GoTo Err_Proc


   cScreen.CenterForm Me

   If Len(ImportFileName) > 0 And Dir(ImportFileName) > vbNullString Then
      optCrLf.Enabled = False
      optComma.Enabled = False
      optSemicolon.Enabled = False
      optTab.Enabled = False
      tMatch = False

      FF = FreeFile
      Open ImportFileName For Input As #FF
      Line Input #FF, tString
      Close #FF
      
      If Len(tString) > 1 Then
         optCrLf.Enabled = True
         tMatch = True
      End If
      If InStr(tString, vbTab) > 0 Then
         optTab.Enabled = True
         tMatch = True
      End If
      If InStr(tString, ";") > 0 Then
         optSemicolon.Enabled = True
         tMatch = True
      End If
      If InStr(tString, ",") > 0 Then
         optComma.Enabled = True
         tMatch = True
      End If
      
      cmdOK.Enabled = tMatch
   End If

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmDelimiters", "Form_Load"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set frmDelimiters = Nothing

End Sub

