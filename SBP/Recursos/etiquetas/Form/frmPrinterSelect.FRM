VERSION 5.00
Begin VB.Form frmPrinterSetUp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Options"
   ClientHeight    =   4380
   ClientLeft      =   3285
   ClientTop       =   1590
   ClientWidth     =   6510
   Icon            =   "frmPrinterSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4380
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   Begin PrintLabels.chameleonButton Command1 
      Height          =   465
      Left            =   5220
      TabIndex        =   16
      Top             =   2355
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "Print"
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
      MICON           =   "frmPrinterSelect.frx":030A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   315
      Left            =   1215
      TabIndex        =   0
      Top             =   240
      Width           =   4845
   End
   Begin PrintLabels.chameleonButton Command2 
      Cancel          =   -1  'True
      Height          =   465
      Left            =   5220
      TabIndex        =   17
      Top             =   2850
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   820
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
      MICON           =   "frmPrinterSelect.frx":0326
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrintLabels.Frame3D frmOptions 
      Height          =   900
      Left            =   210
      Top             =   810
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   1588
      BorderType      =   1
      BevelWidth      =   3
      BevelInner      =   0
      Caption3D       =   1
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
      MouseIcon       =   "frmPrinterSelect.frx":0342
      Picture         =   "frmPrinterSelect.frx":035E
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
      Caption         =   "Options"
      UseMnemonic     =   0   'False
      Begin VB.OptionButton OptCase 
         Caption         =   "Print in Upper/Lower Case"
         Height          =   225
         Index           =   1
         Left            =   165
         TabIndex        =   2
         Top             =   570
         Width           =   2670
      End
      Begin VB.OptionButton OptCase 
         Caption         =   "Print in Upper Case"
         Height          =   225
         Index           =   0
         Left            =   165
         TabIndex        =   1
         Top             =   315
         Value           =   -1  'True
         Width           =   2670
      End
   End
   Begin PrintLabels.Frame3D FraChoices 
      Height          =   2235
      Left            =   210
      Top             =   1785
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3942
      BorderType      =   1
      BevelWidth      =   3
      BevelInner      =   0
      Caption3D       =   1
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
      MouseIcon       =   "frmPrinterSelect.frx":037A
      Picture         =   "frmPrinterSelect.frx":0396
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
      Caption         =   "Report Selection"
      UseMnemonic     =   0   'False
      Begin VB.TextBox txtXTimes 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Text            =   "0"
         Top             =   1125
         Width           =   480
      End
      Begin VB.OptionButton optReport 
         Caption         =   "Print All"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   2685
      End
      Begin VB.OptionButton optReport 
         Caption         =   "Print a Single Page (Starting here)"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   4
         Top             =   555
         Width           =   2685
      End
      Begin VB.OptionButton optReport 
         Caption         =   "Print Selected Zip Codes"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   5
         Top             =   795
         Width           =   2685
      End
      Begin VB.OptionButton optReport 
         Caption         =   "Print Current Label"
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   8
         Top             =   1380
         Width           =   1845
      End
      Begin VB.OptionButton optReport 
         Caption         =   "Print All - Names Only"
         Height          =   255
         Index           =   4
         Left            =   135
         TabIndex        =   9
         Top             =   1710
         Width           =   2685
      End
      Begin VB.OptionButton optReport 
         Caption         =   "Print Zip Codes - Names Only"
         Height          =   255
         Index           =   5
         Left            =   135
         TabIndex        =   10
         Top             =   1950
         Width           =   2685
      End
      Begin VB.OptionButton optReport 
         Caption         =   "Print Current Label"
         Height          =   255
         Index           =   6
         Left            =   150
         TabIndex        =   6
         Top             =   1125
         Width           =   1845
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   15
         X2              =   2985
         Y1              =   1650
         Y2              =   1650
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   15
         X2              =   2980
         Y1              =   1665
         Y2              =   1665
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   15
         X2              =   2980
         Y1              =   1095
         Y2              =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   15
         X2              =   2985
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label3 
         Caption         =   "times."
         Height          =   210
         Left            =   2340
         TabIndex        =   18
         Top             =   1155
         Width           =   525
      End
   End
   Begin PrintLabels.Frame3D fraLocation 
      Height          =   1185
      Left            =   3285
      Top             =   810
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   2090
      BorderType      =   1
      BevelWidth      =   3
      BevelInner      =   0
      Caption3D       =   1
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
      MouseIcon       =   "frmPrinterSelect.frx":03B2
      Picture         =   "frmPrinterSelect.frx":03CE
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
      Caption         =   "Starting Location"
      UseMnemonic     =   0   'False
      Begin VB.ComboBox txt_Down 
         Height          =   315
         Left            =   2295
         Style           =   1  'Simple Combo
         TabIndex        =   12
         Text            =   "1"
         Top             =   705
         Width           =   465
      End
      Begin VB.ComboBox txt_Across 
         Height          =   315
         Left            =   2295
         Style           =   1  'Simple Combo
         TabIndex        =   11
         Text            =   "1"
         Top             =   345
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   " Begin Printing at: "
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   375
         Width           =   1290
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " Row: "
         Height          =   195
         Index           =   5
         Left            =   1800
         TabIndex        =   20
         Top             =   750
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " Column: "
         Height          =   195
         Index           =   4
         Left            =   1605
         TabIndex        =   19
         Top             =   390
         Width           =   660
      End
   End
   Begin PrintLabels.Frame3D fraCopies 
      Height          =   1350
      Left            =   3285
      Top             =   2115
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   2381
      BorderType      =   1
      BevelWidth      =   3
      BevelInner      =   0
      Caption3D       =   1
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
      MouseIcon       =   "frmPrinterSelect.frx":03EA
      Picture         =   "frmPrinterSelect.frx":0406
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
      Caption         =   "Copies"
      UseMnemonic     =   0   'False
      Begin VB.VScrollBar VScroll 
         Height          =   390
         Left            =   1065
         Max             =   -1
         Min             =   -999
         TabIndex        =   14
         Top             =   285
         Value           =   -1
         Width           =   270
      End
      Begin VB.TextBox txtCopies 
         Height          =   285
         Left            =   405
         TabIndex        =   13
         Text            =   "1"
         Top             =   345
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   450
         Index           =   4
         Left            =   120
         Picture         =   "frmPrinterSelect.frx":0422
         Top             =   795
         Width           =   1470
      End
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Printer:"
      Height          =   375
      Index           =   0
      Left            =   255
      TabIndex        =   15
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmPrinterSetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/**********************************
'/* Author: Morgan Haueisen
'/* Copyright (c) 1997-2002
'/**********************************
Option Explicit
Const MaxCopies          As Long = 999
Public ReportType        As Long
Private PrinterName      As String

Private Sub cmbPrinter_Click()

  Dim X As Printer

   For Each X In Printers
      If X.DeviceName = cmbPrinter Then
         'txtDriver = X.DriverName
         'txtPort = X.Port
         Set Printer = X
      End If
   Next X

End Sub

Private Sub Command1_Click()

   On Error Resume Next

   If Not IsNumeric(txtCopies) Then
      txtCopies = 1
   Else
      If Val(txtCopies) < 1 Then txtCopies = 1
      Printer.Copies = CInt(txtCopies)
   End If
   On Error GoTo 0

   SaveSetting App.Title, "User", "DefaultPrinter", cmbPrinter
   Me.Hide
   DoEvents

   If Val(txtXTimes) = 0 Then txtXTimes = 1

   Select Case ReportType
   Case 0 '/* All
      Call PrintAllLabels(True, txt_Across, txt_Down)
   Case 1 '/* Page
      Call PrintAllLabels(False, txt_Across, txt_Down)
   Case 2 '/* zip codes
      With frmPrintZip
         .NamesOnly = False
         .txt_Across = txt_Across
         .txt_Down = txt_Down
         .Show vbModal
      End With
   Case 3 '/* single label
      Call PrintOneX(txt_Across, txt_Down, 1)
   Case 4 '/* Print All - Names Only
      Call PrintNamesOnly(True, txt_Across, txt_Down)
   Case 5 '/* Print Zip Codes - Names Only
      With frmPrintZip
         .NamesOnly = True
         .txt_Across = txt_Across
         .txt_Down = txt_Down
         .Show vbModal
      End With
   Case 6
      Call PrintOneX(txt_Across, txt_Down, txtXTimes)
   Case 7
      frmDataGrid.PrintGridLabels txt_Across, txt_Down
   End Select

   Unload Me

End Sub

Private Sub Command2_Click()

   '    QuitCommand = True

   '    SaveSetting App.Title, "Options", "Printer", PrinterName
   '    SaveSetting App.Title, "Options", "Copies", PrintCopies

   Unload Me

End Sub

Private Sub Form_Load()

   Dim X As Printer
   Dim i As Long
   Dim Index As Long

   On Error GoTo Err_Proc


   cScreen.CenterForm Me

   PrinterName = GetSetting(App.Title, "User", "DefaultPrinter", "None")
   Index = -1

   For Each X In Printers
      cmbPrinter.AddItem X.DeviceName
      If X.DeviceName = PrinterName Then Index = cmbPrinter.NewIndex
   Next X

   If Index = -1 Then
      For i = 0 To cmbPrinter.ListCount - 1
         If cmbPrinter.List(i) = Printer.DeviceName Then
            cmbPrinter.ListIndex = i
         End If
      Next i
   Else
      cmbPrinter.ListIndex = Index
   End If

   PrintCaseType = vbUpperCase

   For i = 1 To NoAcross
      txt_Across.AddItem i
   Next i

   For i = 1 To NoDown
      txt_Down.AddItem i
   Next i
   txtXTimes = NoDown * NoAcross

   VScroll.Max = -MaxCopies
   VScroll.Min = -1
   VScroll.Value = -1
   ReportType = 0

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmPrinterSet", "Form_Load"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set frmPrinterSetUp = Nothing

End Sub

Private Sub OptCase_Click(Index As Integer)

   If OptCase(0).Value Then
      PrintCaseType = vbUpperCase
   Else
      PrintCaseType = vbProperCase
   End If

End Sub

Private Sub optReport_Click(Index As Integer)

   ReportType = Index

End Sub

Private Sub txt_Across_KeyPress(KeyAscii As Integer)

   cValidate.AutoMatch txt_Across, KeyAscii

End Sub

Private Sub txt_Down_KeyPress(KeyAscii As Integer)

   cValidate.AutoMatch txt_Down, KeyAscii

End Sub

Private Sub txtCopies_Change()

   If Val(txtCopies) > MaxCopies Then
      txtCopies = Format$(MaxCopies)
   End If
   VScroll.Value = -Val(txtCopies)

End Sub

Private Sub txtCopies_GotFocus()

   txtGotFocus txtCopies

End Sub

Private Sub txtCopies_KeyPress(KeyAscii As Integer)

   On Error GoTo Err_Proc

   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
   If KeyAscii = 13 Then
      SendKeys "{TAB}"
      KeyAscii = False
   End If
   cValidate.IsNumber txtCopies, KeyAscii, False

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmPrinterSet", "txtCopies_KeyPress"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub txtXTimes_KeyDown(KeyCode As Integer, Shift As Integer)

   On Error GoTo Err_Proc

   Select Case KeyCode
   Case 38  '/* "+"
      txtXTimes = txtXTimes + 1
   Case 40  '/* "-"
      txtXTimes = txtXTimes - 1
      If txtXTimes < 1 Then txtXTimes = 1
   End Select
   KeyCode = False

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmPrinterSet", "txtXTimes_KeyDown"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub txtXTimes_KeyPress(KeyAscii As Integer)

   On Error GoTo Err_Proc

   Select Case KeyAscii
   Case 43 '/* "+"
      txtXTimes = txtXTimes + 1
   Case 45 '/* "-"
      txtXTimes = txtXTimes - 1
      If txtXTimes < 1 Then txtXTimes = 1
   End Select
   KeyAscii = False

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmPrinterSet", "txtXTimes_KeyPress"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub VScroll_Change()

   txtCopies = Abs(VScroll.Value)

End Sub

