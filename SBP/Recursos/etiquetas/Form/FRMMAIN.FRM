VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Label Printer"
   ClientHeight    =   4275
   ClientLeft      =   2400
   ClientTop       =   3930
   ClientWidth     =   7065
   Icon            =   "FRMMAIN.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4275
   ScaleWidth      =   7065
   Begin PrintLabels.chameleonButton cmd_Paste 
      Height          =   795
      Left            =   6000
      TabIndex        =   31
      Top             =   2520
      Visible         =   0   'False
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   1402
      BTYPE           =   3
      TX              =   "Paste"
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
      MICON           =   "FRMMAIN.frx":030A
      PICN            =   "FRMMAIN.frx":0326
      PICH            =   "FRMMAIN.frx":07EC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox PicStatus 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   0
      ScaleHeight     =   1125
      ScaleWidth      =   7065
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3150
      Width           =   7065
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Scheme:"
         ForeColor       =   &H80000014&
         Height          =   270
         Left            =   0
         TabIndex        =   22
         Top             =   30
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Font:"
         ForeColor       =   &H80000014&
         Height          =   270
         Left            =   0
         TabIndex        =   21
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Open File:"
         ForeColor       =   &H80000014&
         Height          =   270
         Left            =   0
         TabIndex        =   20
         Top             =   555
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Paste File:"
         ForeColor       =   &H80000014&
         Height          =   270
         Left            =   0
         TabIndex        =   19
         Top             =   825
         Width           =   975
      End
      Begin VB.Label lblPrintScheme 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000014&
         Height          =   195
         Left            =   1080
         TabIndex        =   18
         Top             =   30
         Width           =   45
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
         ForeColor       =   &H80000014&
         Height          =   240
         Left            =   1080
         TabIndex        =   17
         Top             =   300
         Width           =   45
      End
      Begin VB.Label lbl_Paste 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000014&
         Height          =   195
         Left            =   1080
         TabIndex        =   16
         Top             =   825
         Width           =   45
      End
      Begin VB.Label lblFileOpen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000014&
         Height          =   195
         Left            =   1080
         TabIndex        =   15
         Top             =   555
         Width           =   45
      End
   End
   Begin VB.PictureBox picDataControl 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   390
      Left            =   900
      ScaleHeight     =   330
      ScaleWidth      =   2160
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2790
      Width           =   2220
      Begin VB.CommandButton cmdDataControl 
         Height          =   330
         Index           =   3
         Left            =   1875
         Picture         =   "FRMMAIN.frx":0CB2
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   0
         Width           =   285
      End
      Begin VB.CommandButton cmdDataControl 
         Height          =   330
         Index           =   2
         Left            =   1590
         Picture         =   "FRMMAIN.frx":0D28
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   0
         Width           =   285
      End
      Begin VB.CommandButton cmdDataControl 
         Height          =   330
         Index           =   1
         Left            =   285
         Picture         =   "FRMMAIN.frx":0D9A
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   0
         Width           =   285
      End
      Begin VB.CommandButton cmdDataControl 
         Height          =   330
         Index           =   0
         Left            =   0
         Picture         =   "FRMMAIN.frx":0E0B
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Prev  |  Next"
         ForeColor       =   &H80000017&
         Height          =   330
         Left            =   555
         TabIndex        =   28
         Top             =   45
         Width           =   1050
      End
   End
   Begin VB.PictureBox PicPrinting 
      BackColor       =   &H00A2654A&
      Height          =   3495
      Left            =   8790
      ScaleHeight     =   3435
      ScaleWidth      =   6090
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3600
      Width           =   6150
      Begin VB.CommandButton cmdQuitPrinting 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   405
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   2775
         Width           =   990
      End
      Begin VB.TextBox Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A2654A&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   705
         Left            =   1890
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "FRMMAIN.frx":0E80
         Top             =   225
         Width           =   3735
      End
      Begin VB.TextBox lblPrintUpdate 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2370
         Left            =   1965
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "FRMMAIN.frx":0E98
         Top             =   960
         Width           =   3735
      End
      Begin VB.Image ImpPrinting 
         Height          =   3450
         Left            =   0
         Picture         =   "FRMMAIN.frx":0EB1
         Top             =   0
         Width           =   1650
      End
   End
   Begin VB.PictureBox pnl_labels 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   2400
      Left            =   690
      ScaleHeight     =   2400
      ScaleWidth      =   5190
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   885
      Visible         =   0   'False
      Width           =   5190
      Begin PrintLabels.chameleonButton cmd_Delete 
         Height          =   375
         Left            =   3960
         TabIndex        =   30
         Top             =   1920
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   661
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
         MICON           =   "FRMMAIN.frx":7793
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PrintLabels.chameleonButton cmd_AddNew 
         Height          =   375
         Left            =   2760
         TabIndex        =   29
         Top             =   1920
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Add New"
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
         MICON           =   "FRMMAIN.frx":77AF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.PictureBox pnl_HideLabels 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1755
         Left            =   0
         ScaleHeight     =   1755
         ScaleWidth      =   6000
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   6000
      End
      Begin VB.TextBox txt_line 
         Appearance      =   0  'Flat
         DataField       =   "ZIPCODE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1320
         Width           =   1605
      End
      Begin VB.TextBox txt_line 
         Appearance      =   0  'Flat
         DataField       =   "LINE4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   180
         MaxLength       =   30
         TabIndex        =   11
         Top             =   1320
         Width           =   3195
      End
      Begin VB.TextBox txt_line 
         Appearance      =   0  'Flat
         DataField       =   "LINE3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   180
         MaxLength       =   30
         TabIndex        =   10
         Top             =   960
         Width           =   4785
      End
      Begin VB.TextBox txt_line 
         Appearance      =   0  'Flat
         DataField       =   "LINE2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   180
         MaxLength       =   30
         TabIndex        =   9
         Top             =   600
         Width           =   4785
      End
      Begin VB.TextBox txt_line 
         Appearance      =   0  'Flat
         DataField       =   "LINE1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   180
         MaxLength       =   30
         TabIndex        =   8
         Top             =   240
         Width           =   4785
      End
      Begin VB.Label lblRecPos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   165
         TabIndex        =   1
         Top             =   -15
         Width           =   45
      End
      Begin VB.Label lbl_count 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4845
         TabIndex        =   13
         Top             =   0
         Width           =   75
      End
   End
   Begin PrintLabels.Frame3D fraToolbar 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   873
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
      MouseIcon       =   "FRMMAIN.frx":77CB
      Picture         =   "FRMMAIN.frx":77E7
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
      Begin PrintLabels.chameleonButton cmdQuit 
         Height          =   405
         Left            =   5625
         TabIndex        =   32
         ToolTipText     =   "Exit"
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
         BTYPE           =   9
         TX              =   ""
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
         FOCUSR          =   0   'False
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   65280
         MPTR            =   1
         MICON           =   "FRMMAIN.frx":7803
         PICN            =   "FRMMAIN.frx":781F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PrintLabels.chameleonButton cmdNew 
         Height          =   405
         Left            =   120
         TabIndex        =   33
         ToolTipText     =   "New"
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
         BTYPE           =   9
         TX              =   ""
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
         FOCUSR          =   0   'False
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   255
         MPTR            =   1
         MICON           =   "FRMMAIN.frx":7CE5
         PICN            =   "FRMMAIN.frx":7D01
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PrintLabels.chameleonButton cmdOpen 
         Height          =   405
         Left            =   570
         TabIndex        =   34
         ToolTipText     =   "Open"
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
         BTYPE           =   9
         TX              =   ""
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
         FOCUSR          =   0   'False
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   65280
         MPTR            =   1
         MICON           =   "FRMMAIN.frx":81C7
         PICN            =   "FRMMAIN.frx":81E3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PrintLabels.chameleonButton cmdPrint 
         Height          =   405
         Left            =   1035
         TabIndex        =   35
         ToolTipText     =   "Print"
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
         BTYPE           =   9
         TX              =   ""
         ENAB            =   0   'False
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
         FOCUSR          =   0   'False
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   255
         MPTR            =   1
         MICON           =   "FRMMAIN.frx":86A9
         PICN            =   "FRMMAIN.frx":86C5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PrintLabels.chameleonButton cmdView 
         Height          =   405
         Left            =   1485
         TabIndex        =   36
         ToolTipText     =   "Grid"
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
         BTYPE           =   9
         TX              =   ""
         ENAB            =   0   'False
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
         FOCUSR          =   0   'False
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   255
         MPTR            =   1
         MICON           =   "FRMMAIN.frx":8C93
         PICN            =   "FRMMAIN.frx":8CAF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PrintLabels.chameleonButton cmdFind 
         Height          =   405
         Left            =   1950
         TabIndex        =   37
         ToolTipText     =   "Find"
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
         BTYPE           =   9
         TX              =   ""
         ENAB            =   0   'False
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
         FOCUSR          =   0   'False
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   255
         MPTR            =   1
         MICON           =   "FRMMAIN.frx":91B1
         PICN            =   "FRMMAIN.frx":91CD
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
         Height          =   405
         Left            =   2565
         TabIndex        =   38
         ToolTipText     =   "Fonts"
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
         BTYPE           =   9
         TX              =   ""
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
         FOCUSR          =   0   'False
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   255
         MPTR            =   1
         MICON           =   "FRMMAIN.frx":9693
         PICN            =   "FRMMAIN.frx":96AF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PrintLabels.chameleonButton cmdSchemes 
         Height          =   405
         Left            =   3030
         TabIndex        =   39
         ToolTipText     =   "Schemes"
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
         BTYPE           =   9
         TX              =   ""
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
         FOCUSR          =   0   'False
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   255
         MPTR            =   1
         MICON           =   "FRMMAIN.frx":9AC1
         PICN            =   "FRMMAIN.frx":9ADD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Label Printer"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1575
      TabIndex        =   0
      Top             =   1785
      Width           =   3240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Open Paste &File"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImport 
         Caption         =   "&Import *.txt"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export *.txt"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuLastOpenedSep 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLastOpened 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLastOpened 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLastOpened 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLastOpened 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Tools"
      Begin VB.Menu mnuViewAll 
         Caption         =   "&View All"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFont 
         Caption         =   "&Change Font"
      End
      Begin VB.Menu mnuSpacing 
         Caption         =   "&Set Label Spacing"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cFile      As New clsFileOpenSave

Private Sub AddLastOpened(ByVal PathName As String)

  Dim i                As Long
  Dim n                As Long
  Dim LastOpen(0 To 4) As String

   On Error GoTo Err_Proc


   If PathName = vbNullString Then Exit Sub

   LastOpen(0) = PathName
   n = 0
   For i = 0 To 3
      If PathName <> mnuLastOpened(i).Tag Then
         If Dir(PathName) > vbNullString Then
            n = n + 1
            LastOpen(n) = mnuLastOpened(i).Tag
            mnuLastOpenedSep.Visible = True
         End If
      End If
   Next i

   For i = 0 To 3
      PathName = LastOpen(i)
      If PathName > vbNullString Then
         If Len(PathName) > 35 Then
            n = InStr(Len(PathName) - 35, PathName, "\")
            mnuLastOpened(i).Caption = CStr(i + 1) & ". " & Left$(PathName, 3) & ".." & Mid$(PathName, n)
         Else
            mnuLastOpened(i).Caption = CStr(i + 1) & ". " & PathName
         End If
         mnuLastOpened(i).Tag = PathName
         mnuLastOpened(i).Visible = True
      Else
         mnuLastOpened(i).Visible = False
      End If
   Next i

   Erase LastOpen

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "AddLastOpened"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub cmd_AddNew_Click()

   On Error GoTo Err_Proc

   Call SaveData

   txt_line(0).Visible = True
   txt_line(1).Visible = True
   txt_line(2).Visible = True
   txt_line(3).Visible = True
   txt_line(4).Visible = True
   pnl_HideLabels.Visible = False

   txt_line(0).Text = vbNullString
   txt_line(1).Text = vbNullString
   txt_line(2).Text = vbNullString
   txt_line(3).Text = vbNullString
   txt_line(4).Text = vbNullString

   txt_line(0).SetFocus

   AddingNew = True
   picDataControl.Enabled = True

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "cmd_AddNew_Click"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub cmd_Delete_Click()

  Dim BkMark As Variant

   On Local Error Resume Next

   BkMark = ActiveRS.Bookmark
   ActiveRS.Delete
   ActiveRS.Requery
   ActiveRS.Bookmark = BkMark
   cPlay.PlaySoundResource 1001, , True

   Call UpdateFields

End Sub

Private Sub cmd_Paste_Click()

   On Error Resume Next
   If Not ADOFindFirst(PasteRS, "[LINE1] = '" & txt_line(0) & "' AND [LINE2] = '" & txt_line(1) & "' AND [ZIPCODE] = '" & txt_line(4) & "'") Then
      PasteRS.AddNew
   End If
   PasteRS!Line1 = txt_line(0)
   PasteRS!Line2 = txt_line(1)
   PasteRS!Line3 = txt_line(2)
   PasteRS!Line4 = txt_line(3)
   PasteRS!ZipCode = txt_line(4)
   PasteRS.Update
   On Error GoTo 0
   cPlay.PlaySoundResource 1001, , True

End Sub

Private Sub cmdDataControl_Click(Index As Integer)

   On Error GoTo Err_Proc

   Call SaveData

   Select Case Index
   Case 0
      ActiveRS.MoveFirst
   Case 1
      ActiveRS.MovePrevious
      If ActiveRS.BOF Then ActiveRS.MoveFirst
   Case 2
      ActiveRS.MoveNext
      If ActiveRS.EOF Then ActiveRS.MoveLast
   Case Else
      ActiveRS.MoveLast
   End Select

   Call UpdateFields

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "cmdDataControl_Click"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub cmdFind_Click()

   Dim sFilterStr As String
   Dim FieldName  As String
   Dim FilterType As Integer
   Dim i          As Byte
   Dim FindFirst  As Boolean
   Dim MatchFound As Boolean


   Call SaveData

   On Error GoTo FilterErr

   With frmFilterOptions

      .cmdOK.Caption = "FindFirst"
      .cmdFindNext.Visible = True
      .Move cmdFind.Left + 100, cmdFind.Top + cmdFind.Height + 350
      .Show vbModal

      FindFirst = .FindFirst
      sFilterStr = Trim$(.Text1)
      For i = 0 To 4
         If .OptField(i).Value Then FieldName = .OptField(i).Tag
      Next i
      For i = 0 To 2
         If .optSortOption(i).Value Then FilterType = i
      Next i
   End With
   'Unload frmFilterOptions

   If Len(sFilterStr) > 0 Then
      Select Case FilterType
      Case 0
         sFilterStr = "[" & FieldName & "] LIKE '" & sFilterStr & "%'"
      Case 1
         sFilterStr = "[" & FieldName & "] LIKE '%" & sFilterStr & "%'"
      Case Else
         sFilterStr = "[" & FieldName & "]='" & sFilterStr & "'"
      End Select

      If FindFirst Then
         MatchFound = ADOFindFirst(ActiveRS, sFilterStr)
      Else
         MatchFound = ADOFindNext(ActiveRS, sFilterStr)
      End If

      If MatchFound Then
         Call UpdateFields
      Else
         MsgBox "No match found", vbInformation
      End If
   End If

Exit Sub


FilterErr:
   Screen.MousePointer = vbDefault
   MsgBox "Error:" & Err.Number & " " & Err.Description

End Sub

Private Sub cmdFonts_Click()

   On Error GoTo Err_Proc

   With frmFonts
      .Font_Style = dFontStyle
      .Font_Size = dFontSize
      .Font_Name = dFontName
      .Font_Underline = dFontUnderline
      .Font_StrikeThru = dFontStrikeThru

      .Show vbModal
      If Not QuitCommand Then
         dFontStyle = .Font_Style
         dFontSize = .Font_Size
         dFontName = .Font_Name
         dFontUnderline = .Font_Underline
         dFontStrikeThru = .Font_StrikeThru
      End If
   End With
   Unload frmFonts
   QuitCommand = False
   lblPrintScheme = dPrintScheme
   lblFont = dFontName & ", " & CStr(dFontSize)
   lblFont.Font = dFontName

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "cmdFonts_Click"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub cmdNew_Click()

  Dim Result    As String
  Dim tFileName As String
  Dim tPath     As String


   On Error GoTo Err_Proc
   '/* Open a new MDL FIle */

   '/* Save any changes */
   Call SaveData
   '/* Get New MDL File Name*/
   cFile.VBGetOpenFileName tFileName, , False, , , True, "Labels(*.mdl)|*.mdl", , dbPath, "Create a New Label File", , Me.hwnd
   If tFileName = vbNullString Then Exit Sub

   tPath = cFile.RetOnlyPath(tFileName)
   tFileName = cFile.RetOnlyFilename(tFileName)
   If LCase$(Right$(tFileName, 4)) <> ".mdl" Then tFileName = tFileName & ".mdl"

   If Dir$(tPath & tFileName) > "" Then
      MsgBox "File " & tFileName & " already exists", vbExclamation
      Exit Sub
   End If

   Call CreateMDL(tPath & tFileName)
   Call OpenDataConnection(tPath & tFileName)

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "cmdNew_Click"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub cmdOpen_Click()

   Call OpenLabelFile

End Sub

Private Sub cmdPrint_Click()

   Call SaveData

   frmPrinterSetUp.Show vbModal

End Sub

Private Sub cmdQuit_Click()

   Unload Me

End Sub

Private Sub cmdQuitPrinting_Click()

   PrintingFlag = False

End Sub

Private Sub cmdSchemes_Click()

   frmLabelPlacement.Show vbModal

End Sub

Private Sub cmdView_Click()

   Call SaveData

   frmDataGrid.Show vbModal
   Call UpdateFields

End Sub

Private Sub Form_Initialize()

   Call ManifestWrite

End Sub

Private Sub Form_Load()

  Dim FF      As Long
  Dim i       As Long
  Dim tString As String
  Dim Mydb    As ADODB.Connection
  Dim MySet   As ADODB.Recordset
  Dim c3D     As cls3Dcontrol

   On Error GoTo Err_Proc

   AppPath = App.Path
   If Right$(AppPath, 1) <> "\" Then AppPath = AppPath & "\"

   '/* Get last 4 opened files */
   For i = 3 To 0 Step -1
      tString = GetSetting(App.Title, "User", "LastOpen" & CStr(i), vbNullString)
      If Len(tString) And Dir(tString) > vbNullString Then
         Call AddLastOpened(tString)
      End If
   Next i
   dbPath = GetSetting(App.Title, "User", "dbPath", AppPath)

   '/* Set display defaults */
   Me.Height = 5610
   Me.Width = 6855
   Me.WindowState = vbNormal

   PicPrinting.Move 525, 525
   PicPrinting.Visible = False
   PicPrinting.ZOrder

   mnuPaste.Enabled = False
   mnuImport.Enabled = False
   mnuExport.Enabled = False
   mnuPrint.Enabled = False

   TopMargin = 0.65
   SideMargin = 0.15
   VerPitch = 1
   HorzPitch = 2.73
   NoAcross = 3
   NoDown = 10
   dFontStyle = "regular"
   dFontSize = 10
   dFontName = "Times New Roman"
   dFontUnderline = False
   dFontStrikeThru = False

   '/* Get Last used printing scheme */
   dPrintScheme = GetSetting(App.Title, "User", "PrintScheme", "Default")

   If Dir(AppPath & App.Title & ".dat") = vbNullString Then Call CreateDAT

   Call OpenDB(Mydb, , AppPath & App.Title & ".dat")
   Call OpenRS(MySet, "Select * from Settings", Mydb)
   If Not (MySet.EOF And MySet.BOF) Then
      If ADOFindFirst(MySet, "[Description] = '" & dPrintScheme & "'") Then
         SchemeID = MySet!ID
         TopMargin = MySet!TopMargin
         SideMargin = MySet!SideMargin
         VerPitch = MySet!VPitch
         HorzPitch = MySet!HPitch
         NoAcross = MySet!NoAcross
         NoDown = MySet!NoDown
         dFontSize = MySet!FontSize
         dFontStyle = MySet!FontStyle
         dFontName = MySet!FontName
         dFontUnderline = MySet!FontUnderline
         dFontStrikeThru = MySet!FontStrikethru
         dPrintScheme = MySet!Description
      Else
         dPrintScheme = "Default"
      End If
   End If
   MySet.Close
   Mydb.Close

   If dFontSize = 0 Then dFontSize = 10
   lblFont = dFontName & ", " & CStr(dFontSize)
   lblFont.Font = dFontName
   lblPrintScheme = dPrintScheme

   cScreen.CenterForm Me

   '/* Draw box around address label */
   Set c3D = New cls3Dcontrol
   c3D.ThreeDControl Me, pnl_labels, 1, 4, True
   c3D.ThreeDControl Me, PicStatus
   Set c3D = Nothing

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "Form_Load"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   On Error Resume Next

   If PasteFileOpen Then
      PasteRS.Close
      PasteDB.Close
   End If

End Sub

Private Sub Form_Resize()

   If Me.WindowState = vbMinimized Then Exit Sub

   cmdQuit.Left = Me.Width - cmdQuit.Width - 200

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Dim tfile As String
  Dim i     As Long


   Call SaveData

   For i = 0 To 3
      SaveSetting App.Title, "User", "LastOpen" & CStr(i), mnuLastOpened(i).Tag
   Next i

   On Error Resume Next
   tfile = Dir("*.ldb")
   If tfile > vbNullString Then Kill tfile
   
   '/* Clean up
   Set cFile = Nothing
   Set cScreen = Nothing
   Set cValidate = Nothing
   Set cPlay = Nothing

   Call EndApp
   Set frmMain = Nothing
   '/* End application

End Sub

Private Sub mnuAbout_Click()

   frmAbout.Show

End Sub

Private Sub mnuExit_Click()

   Unload Me

End Sub

Private Sub mnuExport_Click()

  Dim imfile     As String
  Dim LineTxt(5) As String
  Dim tPath      As String
  Dim FF         As Long
  Dim Delimit    As String


   If goApp.SourceDB = vbNullString Then
      MsgBox " You must first open an existing file or create a new one", vbInformation
      Exit Sub
   End If

   pnl_labels.Visible = False

   If Not cFile.VBGetSaveFileName(imfile, , True, "Export(*.txt)|*.txt", , dbPath, "Create an export file name", "txt", Me.hwnd) Then Exit Sub
   If imfile = vbNullString Then Exit Sub

   With frmDelimiters
      .Show vbModal
      If .optComma.Value Then
         Delimit = ","
      ElseIf .optCrLf.Value Then
         Delimit = vbCrLf
      ElseIf .optSemicolon.Value Then
         Delimit = ";"
      ElseIf .optTab.Value Then
         Delimit = vbTab
      End If
   End With
   Unload frmDelimiters
   If QuitCommand Then
      QuitCommand = False
      Exit Sub
   End If

   On Local Error GoTo ExportERROR

   pnl_labels.Visible = True
   mnuPaste.Enabled = True
   mnuImport.Enabled = True
   mnuExport.Enabled = True
   pnl_labels.Refresh

   FF = FreeFile
   Open imfile$ For Output As FF

   ActiveRS.MoveFirst 'Data1.Recordset.MoveFirst
   Do
      If Delimit = vbCrLf Then
         If Len(ActiveRS!Line1) > 0 Then Print #FF, ActiveRS!Line1
         If Len(ActiveRS!Line2) > 0 Then Print #FF, ActiveRS!Line2
         If Len(ActiveRS!Line3) > 0 Then Print #FF, ActiveRS!Line3
         If Len(ActiveRS!Line4) > 0 Then Print #FF, ActiveRS!Line4
         If Len(ActiveRS!ZipCode) > 0 Then Print #FF, ActiveRS!ZipCode & vbCrLf
      Else
         Print #FF, ActiveRS!Line1 & Delimit;
         Print #FF, ActiveRS!Line2 & Delimit;
         Print #FF, ActiveRS!Line3 & Delimit;
         Print #FF, ActiveRS!Line4 & Delimit;
         Print #FF, ActiveRS!ZipCode & vbCrLf
      End If
      ActiveRS.MoveNext
   Loop Until ActiveRS.EOF
   Close #FF

   On Error GoTo 0
   ActiveRS.MoveFirst
   cPlay.PlaySoundResource 1001, , True

Exit Sub


ExportERROR:
   MsgBox Err.Description, vbCritical

End Sub

Private Sub mnuFind_Click()

   cmdFind_Click

End Sub

Private Sub mnuFont_Click()

   Call cmdFonts_Click

End Sub

Private Sub mnuImport_Click()

  Dim imfile     As String
  Dim str_M      As String
  Dim LineTxt(5) As String
  Dim tPath      As String
  Dim tString    As String
  Dim n          As Long
  Dim i          As Long
  Dim FF         As Long
  Dim Delimit    As String


   If goApp.SourceDB = vbNullString Then
      MsgBox " You must first open an existing file or create a new one", vbInformation
      Exit Sub
   End If

   pnl_labels.Visible = False

   If Not cFile.VBGetOpenFileName(imfile, , , , , True, "Import(*.txt;*.csv)|*.txt;*.csv", , dbPath, "Select an Import text file", , Me.hwnd) Then Exit Sub
   If imfile = vbNullString Then Exit Sub

   With frmDelimiters
      .ImportFileName = imfile
      .Show vbModal
      If .optComma.Value Then
         Delimit = ","
      ElseIf .optCrLf.Value Then
         Delimit = vbCrLf
      ElseIf .optSemicolon.Value Then
         Delimit = ";"
      ElseIf .optTab.Value Then
         Delimit = vbTab
      End If
   End With
   Unload frmDelimiters
   If QuitCommand Then
      QuitCommand = False
      Exit Sub
   End If

   On Error GoTo IMPORTERR

   pnl_labels.Visible = True
   mnuPaste.Enabled = True
   mnuImport.Enabled = True
   mnuExport.Enabled = True
   pnl_labels.Refresh

   FF = FreeFile
   Open imfile$ For Input As FF
   If Delimit = vbCrLf Then
      Do
         Line Input #FF, tString
         If EOF(FF) Then Exit Do
         tString = Trim$(tString)
         If tString > vbNullString Then
            n = n + 1
            LineTxt(n) = tString
            If n > 2 And Asc(Right$(LineTxt(n), 1)) > 47 And Asc(Right$(LineTxt(n), 1)) < 58 Then
               For i = Len(LineTxt(n)) To 1 Step -1
                  If Mid$(LineTxt(n), i, 1) = " " Then Exit For
               Next i
               If i > 1 Then tString = Trim$(Mid$(LineTxt(n), i))
               If Len(tString) = 5 Or Len(tString) = 10 Then
                  If i > 1 Then
                     n = n + 1
                     LineTxt(n) = Trim$(Mid$(LineTxt(n - 1), i))
                     LineTxt(n - 1) = RTrim$(Left$(LineTxt(n - 1), i))
                  End If
                  GoSub SaveData
               End If
            End If
         Else
            If n > 2 Then GoSub SaveData
         End If
      Loop
      If n > 2 Then GoSub SaveData

   Else

      Do
         Line Input #FF, tString
         If EOF(FF) Then Exit Do
         tString = Trim$(tString)
         If tString > "" Then
            Erase LineTxt
            For n = 1 To 4
               i = InStr(tString, Delimit)
               If i > 0 Then
                  LineTxt(n) = Mid$(tString, 1, i - 1)
                  tString = Mid$(tString, i + 1)
               End If
            Next n
            LineTxt(5) = tString

            ActiveRS.AddNew
            txt_line(0).Text = LineTxt(1)
            txt_line(1).Text = LineTxt(2)
            txt_line(2).Text = LineTxt(3)
            txt_line(3).Text = LineTxt(4)
            txt_line(4).Text = LineTxt(5)
            ActiveRS.Update
         End If
      Loop
   End If
   Close #FF

   cPlay.PlaySoundResource 1001, , True

EXITIMPORT:
   On Error GoTo 0

Exit Sub


IMPORTERR:
   str_M = "Can not seperate address lines at address number " & Format$(AddNumber&) & Chr$(10) & "Last Read:"
   
   For i = 1 To n - 1
      str_M = str_M & Chr$(10) & LineTxt(i)
   Next i
   MsgBox Err.Description & vbCrLf & str_M, vbInformation, "ERROR"
   Resume EXITIMPORT

SaveData:
   ActiveRS.AddNew
   For i = 1 To n
      If i = n And Asc(Right$(LineTxt(n), 1)) > 47 And Asc(Right$(LineTxt(n), 1)) < 58 Then
         txt_line(4).Text = LineTxt(i)
         txt_line(4).Refresh
      Else
         txt_line(i - 1).Text = LineTxt(i)
         txt_line(i - 1).Refresh
      End If
   Next i
   ActiveRS.Update
   DoEvents
   n = False
   Erase LineTxt
   Return

End Sub

Private Sub mnuLastOpened_Click(Index As Integer)

   Call OpenLabelFile(mnuLastOpened(Index).Tag)

End Sub

Private Sub mnuNew_Click()

   Call cmdNew_Click

End Sub

Private Sub mnuOpen_Click()

   Call OpenLabelFile

End Sub

Private Sub mnuPaste_Click()

  Dim SQLstmt   As String
  Dim tFileName As String


   If Not cFile.VBGetOpenFileName(tFileName, , , , , True, "Labels(*.mdl)|*.mdl", , dbPath, "Open Label Paste File", , Me.hwnd) Then Exit Sub
   If tFileName = vbNullString Then Exit Sub

   If LCase$(goApp.SourceDB) = LCase$(tFileName) Then
      MsgBox "The paste file and the currently open can not be the same", vbCritical
      Exit Sub
   End If

   On Local Error Resume Next
   If PasteFileOpen Then
      PasteRS.Close
      PasteDB.Close
      lbl_Paste = vbNullString
   End If
   PasteFile = True

   '/* Open database and recordset */
   Call OpenDB(PasteDB, , tFileName)
   SQLstmt = "SELECT * FROM LABELS ORDER BY [ZIPCODE]"
   Call OpenRS(PasteRS, SQLstmt, PasteDB)
   PasteFileOpen = True
   cmd_Paste.Visible = PasteFileOpen

   tFileName = cFile.RetOnlyFilename(tFileName)
   lbl_Paste = Left$(tFileName, Len(tFileName) - 4)

End Sub

Private Sub mnuPrint_Click()

   Call cmdPrint_Click

End Sub

Private Sub mnuSpacing_Click()

   frmLabelPlacement.Show vbModal

End Sub

Private Sub mnuViewAll_Click()

   Call cmdView_Click

End Sub

Private Sub OpenDataConnection(ByVal tFileName As String)

  Dim SQLstmt    As String
  Dim dbFileName As String

   On Error GoTo Err_Proc


   dbFileName = cFile.RetOnlyFilename(tFileName)
   dbPath = cFile.RetOnlyPath(tFileName)
   dbPath = IIf(Right$(dbPath, 1) = "\", dbPath, dbPath & "\")
   SaveSetting App.Title, "User", "dbPath", dbPath

   Call InitSettings(dbPath & dbFileName)
   Call AddLastOpened(dbPath & dbFileName)

   SQLstmt = "SELECT LABELS.* FROM LABELS ORDER BY [ZIPCODE],[Line1]"
   Call OpenDB(ActiveDB)
   Call OpenRS(ActiveRS, SQLstmt, ActiveDB)

   lblFileOpen = Left$(dbFileName, Len(dbFileName) - 4)

   mnuPaste.Enabled = True
   mnuImport.Enabled = True
   mnuExport.Enabled = True
   mnuPrint.Enabled = True
   cmdPrint.Enabled = True
   cmdView.Enabled = True
   cmdFind.Enabled = True
   pnl_labels.Visible = True

   Call UpdateFields

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "OpenDataConnection"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub OpenLabelFile(Optional ByVal tFileName As String = vbNullString)

   On Error GoTo Err_Proc

   Call SaveData
   If tFileName = vbNullString Then
      cFile.VBGetOpenFileName tFileName, , , , , True, "Labels(*.mdl)|*.mdl", , dbPath, "Open Label File", , Me.hwnd
      If tFileName = vbNullString Then Exit Sub
   End If

   Call OpenDataConnection(tFileName)

   lbl_count = CStr(ActiveRS.RecordCount)

   If PasteFileOpen Then
      PasteRS.Close
      PasteDB.Close
      PasteFileOpen = False
      cmd_Paste.Visible = False
      mnuPaste.Enabled = False
      lbl_Paste = vbNullString
   End If

   Call UpdateFields

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "OpenLabelFile"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub SaveData()

   On Local Error Resume Next

   If LenB(txt_line(0)) = 0 Then
      AddingNew = False
      Exit Sub
   End If

   If AddingNew Then ActiveRS.AddNew

   ActiveRS!Line1 = txt_line(0).Text
   ActiveRS!Line2 = txt_line(1).Text
   ActiveRS!Line3 = txt_line(2).Text
   ActiveRS!Line4 = txt_line(3).Text
   ActiveRS!ZipCode = txt_line(4).Text
   ActiveRS.Update
   AddingNew = False

End Sub

Private Sub txt_line_Change(Index As Integer)

  Dim X As Long

   On Error GoTo Err_Proc


   cValidate.FcaseTextBox txt_line(Index)

   X = txt_line(Index).SelStart
   If Len(txt_line(Index)) > 3 Then
      If Mid$(txt_line(Index), Len(txt_line(Index)) - 2, 1) = " " Then
         txt_line(Index) = Left$(txt_line(Index), Len(txt_line(Index)) - 1) & UCase$(Right$(txt_line(Index), 1))
         
      End If
   End If
   txt_line(Index).SelStart = X

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "txt_line_Change"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub txt_line_KeyPress(Index As Integer, KeyAscii As Integer)

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{TAB}"
   End If

End Sub

Private Sub txt_line_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

   If Index = 4 Then
      cValidate.MaskedText txt_line(Index), "#####-####"
   End If

End Sub

Private Sub txt_line_LostFocus(Index As Integer)

   On Error GoTo Err_Proc

   txt_line(Index) = Trim$(txt_line(Index))
   If Index = 2 Or Index = 3 Then
      If txt_line(Index) > vbNullString Then
         If Right$(txt_line(Index), 1) = "." Then
            txt_line(Index) = Left$(txt_line(Index), Len(txt_line(Index)) - 1)
         End If
      End If
   End If

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "txt_line_LostFocus"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub UpdateFields()

  Dim rPos As Long

   On Error GoTo Err_Proc

   If ActiveRS.RecordCount = 0 Then
      txt_line(0).Visible = False
      txt_line(1).Visible = False
      txt_line(2).Visible = False
      txt_line(3).Visible = False
      txt_line(4).Visible = False
      cmd_Delete.Visible = False
      pnl_HideLabels.Visible = True
      picDataControl.Enabled = False

      txt_line(0).Text = vbNullString
      txt_line(1).Text = vbNullString
      txt_line(2).Text = vbNullString
      txt_line(3).Text = vbNullString
      txt_line(4).Text = vbNullString

      rPos = 0

   Else
      txt_line(0).Visible = True
      txt_line(1).Visible = True
      txt_line(2).Visible = True
      txt_line(3).Visible = True
      txt_line(4).Visible = True
      cmd_Delete.Visible = True
      pnl_HideLabels.Visible = False
      picDataControl.Enabled = True

      txt_line(0).Text = ActiveRS!Line1 & vbNullString
      txt_line(1).Text = ActiveRS!Line2 & vbNullString
      txt_line(2).Text = ActiveRS!Line3 & vbNullString
      txt_line(3).Text = ActiveRS!Line4 & vbNullString
      txt_line(4).Text = ActiveRS!ZipCode & vbNullString
      txt_line(0).SetFocus

      rPos = ActiveRS.AbsolutePosition

   End If

   lblRecPos = "Record: " & CStr(rPos)
   lbl_count = CStr(ActiveRS.RecordCount)

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "UpdateFields"
   Err.Clear
   Resume Exit_Here

End Sub

