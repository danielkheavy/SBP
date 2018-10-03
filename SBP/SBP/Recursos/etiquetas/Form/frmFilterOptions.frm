VERSION 5.00
Begin VB.Form frmFilterOptions 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3210
   ClientLeft      =   4320
   ClientTop       =   2745
   ClientWidth     =   3795
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmFilterOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   Begin PrintLabels.chameleonButton cmdQuit 
      Height          =   435
      Left            =   540
      TabIndex        =   2
      Top             =   2595
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   767
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
      MICON           =   "frmFilterOptions.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   630
      TabIndex        =   0
      Top             =   2130
      Width           =   3015
   End
   Begin PrintLabels.chameleonButton cmdOK 
      Height          =   435
      Left            =   1500
      TabIndex        =   3
      Top             =   2595
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   767
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
      MICON           =   "frmFilterOptions.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrintLabels.chameleonButton cmdFindNext 
      Height          =   435
      Left            =   2475
      TabIndex        =   4
      Top             =   2595
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   767
      BTYPE           =   3
      TX              =   "Next"
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
      MICON           =   "frmFilterOptions.frx":0044
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
      Height          =   1875
      Left            =   165
      Top             =   120
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   3307
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
      MouseIcon       =   "frmFilterOptions.frx":0060
      Picture         =   "frmFilterOptions.frx":007C
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
      Caption         =   "Field Name"
      UseMnemonic     =   0   'False
      Begin VB.OptionButton OptField 
         Caption         =   "Zip Code"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Tag             =   "zipcode"
         Top             =   1515
         Value           =   -1  'True
         Width           =   1320
      End
      Begin VB.OptionButton OptField 
         Caption         =   "Line 4"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Tag             =   "line4"
         Top             =   1230
         Width           =   1320
      End
      Begin VB.OptionButton OptField 
         Caption         =   "Line 3"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Tag             =   "line3"
         Top             =   930
         Width           =   1320
      End
      Begin VB.OptionButton OptField 
         Caption         =   "Line 2"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Tag             =   "line2"
         Top             =   645
         Width           =   1320
      End
      Begin VB.OptionButton OptField 
         Caption         =   "Line 1"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Tag             =   "line1"
         Top             =   360
         Width           =   1320
      End
   End
   Begin PrintLabels.Frame3D Frame2 
      Height          =   1875
      Left            =   1695
      Top             =   120
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   3307
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
      MouseIcon       =   "frmFilterOptions.frx":0098
      Picture         =   "frmFilterOptions.frx":00B4
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
      Caption         =   "Filter Type"
      UseMnemonic     =   0   'False
      Begin VB.OptionButton optSortOption 
         Caption         =   "Exact Match"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   12
         Top             =   885
         Width           =   1290
      End
      Begin VB.OptionButton optSortOption 
         Caption         =   "Contains"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   11
         Top             =   585
         Value           =   -1  'True
         Width           =   1290
      End
      Begin VB.OptionButton optSortOption 
         Caption         =   "Begins with"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   10
         Top             =   300
         Width           =   1305
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " Find: "
      Height          =   195
      Left            =   165
      TabIndex        =   1
      Top             =   2160
      Width           =   435
   End
End
Attribute VB_Name = "frmFilterOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FindFirst      As Boolean
Public FindNext       As Boolean

Private Sub cmdFindNext_Click()

   FindFirst = False
   FindNext = True
   Me.Hide

End Sub

Private Sub cmdOK_Click()

   FindFirst = True
   FindNext = False
   Me.Hide

End Sub

Private Sub cmdQuit_Click()

   QuitCommand = True
   Me.Hide

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set frmFilterOptions = Nothing

End Sub

