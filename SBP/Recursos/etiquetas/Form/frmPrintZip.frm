VERSION 5.00
Begin VB.Form frmPrintZip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print a ZipCode"
   ClientHeight    =   4980
   ClientLeft      =   3750
   ClientTop       =   1650
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   Begin PrintLabels.chameleonButton cmdPrint 
      Height          =   450
      Left            =   3060
      TabIndex        =   2
      Top             =   3555
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   794
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
      MICON           =   "frmPrintZip.frx":0000
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
      Height          =   4560
      Left            =   225
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   2400
   End
   Begin PrintLabels.chameleonButton cmdQuit 
      Height          =   450
      Left            =   3045
      TabIndex        =   3
      Top             =   4065
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   794
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
      MICON           =   "frmPrintZip.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrintLabels.Frame3D fraBegin 
      Height          =   1110
      Left            =   2730
      Top             =   2340
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1958
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
      MouseIcon       =   "frmPrintZip.frx":0038
      Picture         =   "frmPrintZip.frx":0054
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
      Caption         =   "Begin Printing at:"
      UseMnemonic     =   0   'False
      Begin VB.ComboBox txt_Across 
         Height          =   315
         Left            =   1635
         Style           =   1  'Simple Combo
         TabIndex        =   5
         Text            =   "1"
         Top             =   300
         Width           =   465
      End
      Begin VB.ComboBox txt_Down 
         Height          =   315
         Left            =   1635
         Style           =   1  'Simple Combo
         TabIndex        =   4
         Text            =   "1"
         Top             =   660
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " Row: "
         Height          =   195
         Index           =   5
         Left            =   1140
         TabIndex        =   7
         Top             =   735
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " Column: "
         Height          =   195
         Index           =   4
         Left            =   945
         TabIndex        =   6
         Top             =   375
         Width           =   660
      End
   End
   Begin VB.Label Label3 
      Height          =   1620
      Left            =   2880
      TabIndex        =   1
      Top             =   225
      Width           =   2055
   End
End
Attribute VB_Name = "frmPrintZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Mydb          As ADODB.Connection
Private MySet         As ADODB.Recordset
Public NamesOnly      As Boolean

Private Sub cmdPrint_Click()

   Call PrintZipCodes

End Sub

Private Sub cmdQuit_Click()

   On Local Error Resume Next

   MySet.Close
   Mydb.Close
   Unload Me

End Sub

Private Sub Form_Load()

  Dim SQLstmt As String
  Dim i       As Long

   On Error GoTo Err_Proc


   cScreen.CenterForm Me
   Me.Icon = frmMain.Icon
   Label3.Caption = "Select the zip codes to print." & vbCrLf & vbCrLf & _
            "The trailing number is the number of labels that will be printed for each zip code selected."

   SQLstmt = "SELECT DISTINCTROW First(Labels.ZIPCODE) AS [ZIPCODES], Count(Labels.ZIPCODE) AS NumberOf"
   SQLstmt = SQLstmt & " From Labels"
   SQLstmt = SQLstmt & " GROUP BY Labels.ZIPCODE"
   SQLstmt = SQLstmt & " Having (((Count(Labels.ZIPCODE)) > 1))"
   SQLstmt = SQLstmt & " ORDER BY First(Labels.ZIPCODE);"

   Call OpenDB(Mydb)
   Call OpenRS(MySet, SQLstmt, Mydb)
   If Not (MySet.EOF And MySet.BOF) Then
      Do
         List1.AddItem MySet!Zipcodes & " == " & MySet!NumberOf
         MySet.MoveNext
      Loop Until MySet.EOF
   End If
   MySet.Close

   For i = 1 To NoAcross
      txt_Across.AddItem i
   Next i

   For i = 1 To NoDown
      txt_Down.AddItem i
   Next i

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmPrintZip", "Form_Load"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set frmPrintZip = Nothing

End Sub

Private Sub PrintZipCodes()

  Dim EndOfFile         As Boolean
  Dim Down              As Long
  Dim HoldPlace         As Variant
  Dim hTabPos           As Single
  Dim VTabPos           As Single
  Dim StartCol          As Long
  Dim StartRow          As Long
  Dim SQLstmt           As String
  Dim i                 As Long
  Dim n                 As Long
  Dim WhereCondition    As Boolean


   StartCol = txt_Across
   StartRow = txt_Down

   SQLstmt = "SELECT * From Labels Where"
   For i = 0 To List1.ListCount - 1
      If List1.Selected(i) Then
         n = InStr(List1.List(i), "==")
         If WhereCondition Then
            SQLstmt = SQLstmt & " OR Labels.ZipCode = '" & Left$(List1.List(i), n - 2) & "'"
         Else
            SQLstmt = SQLstmt & " Labels.ZipCode = '" & Left$(List1.List(i), n - 2) & "'"
            WhereCondition = True
         End If
      End If
   Next i
   SQLstmt = SQLstmt & " Order by Labels.ZipCode"

   If WhereCondition = False Then
      Mydb.Close
      Unload Me
      Exit Sub
   End If

   Me.Visible = False
   PrintingFlag = True
   PrintingShow True

   Call OpenRS(MySet, SQLstmt, Mydb)
   On Local Error Resume Next

   Call SetUpPage

   hTabPos = SideMargin
   VTabPos = TopMargin

   Do

      ReDim Line1(NoAcross) As String
      ReDim Line2(NoAcross) As String
      ReDim Line3(NoAcross) As String
      ReDim Line4(NoAcross) As String
      ReDim line5(NoAcross) As String

      For i = StartCol To NoAcross
         PrintingShow , MySet!Line1

         If NamesOnly Then
            Line1(i) = vbNullString
            Line2(i) = FixStringCase(MySet!Line1 & vbNullString, PrintCaseType)
            Line3(i) = vbNullString
            Line4(i) = vbNullString
            MySet.MoveNext
            If MySet.EOF Then
               EndOfFile = True
               Exit For
            End If
         Else
            Line1(i) = FixStringCase(MySet!Line1 & vbNullString, PrintCaseType)
            Line2(i) = FixStringCase(MySet!Line2 & vbNullString, PrintCaseType)
            Line3(i) = FixStringCase(MySet!Line3 & vbNullString, PrintCaseType)
            Line4(i) = FixStringCase(MySet!Line4 & vbNullString, PrintCaseType)
            If Trim$(Line4(i)) = vbNullString Then
               Line3(i) = Trim$(Line3(i))
               Line3(i) = Left$(Line3(i), Len(Line3(i)) - 1) & UCase$(Right$(Line3(i), 1))
            Else
               Line4(i) = Trim$(Line4(i))
               Line4(i) = Left$(Line4(i), Len(Line4(i)) - 1) & UCase$(Right$(Line4(i), 1))
            End If
            line5(i) = MySet!ZipCode
            MySet.MoveNext
            If MySet.EOF Then
               EndOfFile = True
               Exit For
            End If
         End If
      Next i
      GoSub PrintLines

   Loop Until EndOfFile

   Printer.EndDoc

   PrintingShow False
   MySet.Close
   Mydb.Close
   Unload Me

Exit Sub


PrintLines:
   If StartRow > 1 Then
      VTabPos = VerPitch * StartRow
      Printer.CurrentY = VTabPos
      Down = StartRow
      StartRow = 1
   Else
      Down = Down + 1
      Printer.CurrentY = VTabPos
   End If

   DoEvents
   hTabPos = SideMargin
   If StartCol > 1 Then hTabPos = (hTabPos + HorzPitch) * (StartCol - 1)

   If NamesOnly Then
      Printer.Print
      For i = StartCol To NoAcross
         Printer.CurrentX = hTabPos
         Printer.Print Line2(i);
         If i = NoAcross Then Printer.Print Else hTabPos = hTabPos + HorzPitch
      Next i
   Else
      For i = StartCol To NoAcross
         Printer.CurrentX = hTabPos
         Printer.Print Line1(i);
         If i = NoAcross Then Printer.Print Else hTabPos = hTabPos + HorzPitch
      Next i

      DoEvents
      hTabPos = SideMargin
      If StartCol > 1 Then hTabPos = (hTabPos + HorzPitch) * (StartCol - 1)
      For i = StartCol To NoAcross
         Printer.CurrentX = hTabPos
         Printer.Print Line2(i);
         If i = NoAcross Then Printer.Print Else hTabPos = hTabPos + HorzPitch
      Next i

      DoEvents
      hTabPos = SideMargin
      If StartCol > 1 Then hTabPos = (hTabPos + HorzPitch) * (StartCol - 1)
      For i = StartCol To NoAcross
         Printer.CurrentX = hTabPos
         If Line4(i) = vbNullString Then
            n = Len(Line3(i))
            If n > 0 Then
               Line3(i) = Left$(Line3(i), n - 2) & UCase$(Right$(Line3(i), 2))
            End If
            Printer.Print Line3(i); "  " & line5(i);
         Else
            Printer.Print Line3(i);
         End If
         If i = NoAcross Then Printer.Print Else hTabPos = hTabPos + HorzPitch
      Next i

      DoEvents
      hTabPos = SideMargin
      If StartCol > 1 Then hTabPos = (hTabPos + HorzPitch) * (StartCol - 1)
      For i = StartCol To NoAcross
         If Line4(i) > vbNullString Then
            Printer.CurrentX = hTabPos
            n = Len(Line4(i))
            If n > 0 Then
               Line4(i) = Left$(Line4(i), n - 2) & UCase$(Right$(Line4(i), 2))
            End If
            Printer.Print Line4(i); "  "; line5(i);
         End If
         If i = NoAcross Then Printer.Print Else hTabPos = hTabPos + HorzPitch
      Next i
   End If

   If Down = NoDown Then
      Call SetUpNewPage
      Down = False
      VTabPos = TopMargin
   Else
      VTabPos = VTabPos + VerPitch
   End If

   StartCol = 1
   DoEvents
   If PrintingFlag = False Then EndOfFile = True

   Return

End Sub

Private Sub txt_Across_KeyPress(KeyAscii As Integer)

   cValidate.AutoMatch txt_Across, KeyAscii

End Sub

Private Sub txt_Down_KeyPress(KeyAscii As Integer)

   cValidate.AutoMatch txt_Down, KeyAscii

End Sub

