VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmDataGrid 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4845
   ClientLeft      =   1560
   ClientTop       =   2430
   ClientWidth     =   8760
   Icon            =   "frmDataGrid_Sheridan.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleMode       =   0  'User
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid grdDataGrid 
      Align           =   1  'Align Top
      Bindings        =   "frmDataGrid_Sheridan.frx":000C
      Height          =   1800
      Left            =   0
      TabIndex        =   7
      Top             =   630
      Width           =   8760
      _Version        =   196617
      AllowDelete     =   -1  'True
      AllowUpdate     =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   3
      SelectByCell    =   -1  'True
      ForeColorEven   =   -2147483640
      BackColorEven   =   -2147483643
      RowHeight       =   423
      Columns.Count   =   5
      Columns(0).Width=   3200
      Columns(0).Caption=   "LINE1"
      Columns(0).Name =   "LINE1"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "LINE1"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   5292
      Columns(1).Caption=   "LINE2"
      Columns(1).Name =   "LINE2"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "LINE2"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   5292
      Columns(2).Caption=   "LINE3"
      Columns(2).Name =   "LINE3"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "LINE3"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3519
      Columns(3).Caption=   "LINE4"
      Columns(3).Name =   "LINE4"
      Columns(3).CaptionAlignment=   0
      Columns(3).DataField=   "LINE4"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Caption=   "ZIPCODE"
      Columns(4).Name =   "ZIPCODE"
      Columns(4).CaptionAlignment=   0
      Columns(4).DataField=   "ZIPCODE"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      _ExtentX        =   15452
      _ExtentY        =   3175
      _StockProps     =   79
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picButtons 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   8760
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8760
      Begin PrintLabels.chameleonButton cmdRefresh 
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   15
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1085
         BTYPE           =   9
         TX              =   "&Refresh"
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
         MICON           =   "frmDataGrid_Sheridan.frx":0020
         PICN            =   "frmDataGrid_Sheridan.frx":003C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PrintLabels.chameleonButton cmdFilter 
         Height          =   615
         Left            =   1275
         TabIndex        =   2
         Top             =   15
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1085
         BTYPE           =   9
         TX              =   "&Filter"
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
         MCOL            =   0
         MPTR            =   1
         MICON           =   "frmDataGrid_Sheridan.frx":046A
         PICN            =   "frmDataGrid_Sheridan.frx":0486
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PrintLabels.chameleonButton cmdDelete 
         Height          =   615
         Left            =   2430
         TabIndex        =   3
         Top             =   15
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1085
         BTYPE           =   9
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
         MCOL            =   255
         MPTR            =   1
         MICON           =   "frmDataGrid_Sheridan.frx":0A28
         PICN            =   "frmDataGrid_Sheridan.frx":0A44
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PrintLabels.chameleonButton cmdPaste 
         Height          =   615
         Left            =   3585
         TabIndex        =   4
         Top             =   15
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1085
         BTYPE           =   9
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
         MICON           =   "frmDataGrid_Sheridan.frx":0F0A
         PICN            =   "frmDataGrid_Sheridan.frx":0F26
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PrintLabels.chameleonButton cmdPrintSel 
         Height          =   615
         Left            =   4740
         TabIndex        =   5
         Top             =   15
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1085
         BTYPE           =   9
         TX              =   "Print Selected"
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
         MICON           =   "frmDataGrid_Sheridan.frx":13EC
         PICN            =   "frmDataGrid_Sheridan.frx":1408
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PrintLabels.chameleonButton cmdClose 
         Height          =   615
         Left            =   5895
         TabIndex        =   6
         Top             =   15
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1085
         BTYPE           =   9
         TX              =   "&Close"
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
         MICON           =   "frmDataGrid_Sheridan.frx":19D6
         PICN            =   "frmDataGrid_Sheridan.frx":19F2
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   0
      Top             =   2685
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmDataGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub cmdClose_Click()

   Unload Me

End Sub

Private Sub cmdDelete_Click()

   Dim vBkMark As Variant
   Dim i       As Long
   Dim c       As Long



   On Local Error Resume Next
   If grdDataGrid.SelBookmarks.Count > 1 Then
      Do
         vBkMark = grdDataGrid.SelBookmarks(0)
         Data1.Recordset.Bookmark = vBkMark
         Data1.Recordset.Delete
      Loop Until grdDataGrid.SelBookmarks.Count = 0
   Else
      vBkMark = Data1.Recordset.Bookmark
      Data1.Recordset.Delete
   End If

   Data1.Recordset.Requery
   Data1.Refresh
   Data1.Recordset.Bookmark = vBkMark

   DoEvents
   ActiveRS.Requery
   DoEvents

   cPlay.PlaySoundResource 1001, , True
   On Local Error GoTo 0

End Sub

Private Sub cmdFilter_Click()

   Dim FieldName  As String
   Dim FilterType As Long
   Dim i          As Byte
   Dim sFilterStr As String

   On Error GoTo FilterErr

   With frmFilterOptions
      .Move cmdFilter.Left + 100, cmdFilter.Top + cmdFilter.Height + 350
      .Show vbModal
      sFilterStr = Trim$(.Text1)
      For i = 0 To 4
         If .OptField(i).Value Then FieldName = .OptField(i).Tag
      Next i
      For i = 0 To 2
         If .optSortOption(i).Value Then FilterType = i
      Next i
   End With
   Unload frmFilterOptions

   If Len(sFilterStr) > 0 Then
      Select Case FilterType
      Case 0
         sFilterStr = "[" & FieldName & "] LIKE '" & sFilterStr & "%'"
      Case 1
         sFilterStr = "[" & FieldName & "] LIKE '%" & sFilterStr & "%'"
      Case Else
         sFilterStr = "[" & FieldName & "]='" & sFilterStr & "'"
      End Select
      'sFilterStr = InputBox("Enter Filter Expression:" & vbNewLine & "Example format: [zipcode]='30643'")
      Data1.Recordset.Filter = sFilterStr
   End If

Exit Sub


FilterErr:
   Screen.MousePointer = vbDefault
   MsgBox "Error:" & Err.Number & " " & Err.Description

End Sub

Private Sub cmdPaste_Click()

   Dim vBkMark  As Variant
   Dim SQLstmt  As String
   Dim tLine1   As String
   Dim tLine2   As String
   Dim tLine3   As String
   Dim tLine4   As String
   Dim tZipCode As String
   Dim i        As Long


   On Local Error Resume Next

   For i = 0 To (grdDataGrid.SelBookmarks.Count - 1)
      vBkMark = grdDataGrid.SelBookmarks(i)
      tLine1 = grdDataGrid.Columns(1).CellText(vBkMark)
      tLine2 = grdDataGrid.Columns(2).CellText(vBkMark)
      tLine3 = grdDataGrid.Columns(3).CellText(vBkMark)
      tLine4 = grdDataGrid.Columns(4).CellText(vBkMark)
      tZipCode = grdDataGrid.Columns(5).CellText(vBkMark)

      SQLstmt = "[LINE1] = '" & tLine1 & "' AND [LINE2] = '" & tLine2 & "' AND [ZIPCODE] = '" & tZipCode & "'"
      
      If Not ADOFindFirst(PasteRS, SQLstmt) Then
         PasteRS.AddNew
      End If
      PasteRS!Line1 = tLine1
      PasteRS!Line2 = tLine2
      PasteRS!Line3 = tLine3
      PasteRS!Line4 = tLine4
      PasteRS!ZipCode = tZipCode
      PasteRS.Update
   Next i

   cPlay.PlaySoundResource 1001, , True
   On Local Error GoTo 0

End Sub

Private Sub cmdPrintSel_Click()

   On Error GoTo Err_Proc

   If grdDataGrid.SelBookmarks.Count > 0 Then
      With frmPrinterSetUp
         .FraChoices.Visible = False
         .ReportType = 7
         .Show vbModal
      End With
   End If
   QuitCommand = False

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmDataGrid", "cmdPrintSel_Click"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub cmdRefresh_Click()

   On Error GoTo RefErr

   Data1.RecordSource = "select * from labels order by [zipcode]"
   Data1.Refresh

Exit Sub


RefErr:
   MsgBox "Error:" & Err.Number & " " & Err.Description

End Sub

Private Sub Form_Load()

   cScreen.FitScreen Me

   On Error GoTo LoadErr

   ADOdcConnect Data1, ActiveRS.Source, goApp.SourceDB, adOpenDynamic
   ADOdcFindFirst Data1, "[ID]=" & ActiveRS!ID

   ActiveRS.Close
   ActiveDB.Close

   Me.Caption = Data1.RecordSource
   Me.Icon = frmMain.Icon
   cmdPaste.Visible = PasteFileOpen

Exit Sub


LoadErr:
   MsgBox "Error:" & Err.Number & " " & Err.Description
   Unload Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   Call OpenDB(ActiveDB)

   Call OpenRS(ActiveRS, Data1.RecordSource, ActiveDB)
   ADOFindFirst ActiveRS, "[ID]=" & Format$(Data1.Recordset!ID)
   DoEvents

End Sub

Private Sub Form_Resize()

   On Error Resume Next

   If Me.WindowState <> 1 Then
      grdDataGrid.Height = Me.Height - (425 + picButtons.Height)
      'grdDataGrid.Width = Me.Width - 200
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

   On Local Error Resume Next

   Set frmDataGrid = Nothing

End Sub

Private Sub grdDataGrid_DblClick()

   If PasteFileOpen Then

      cmdPaste_Click
   Else
      cmdClose_Click
   End If

End Sub

Private Sub grdDataGrid_HeadClick(ByVal ColIndex As Integer)

   Dim SortCol        As String
   Static SortOrder   As Boolean
   Static LastColName As String

   On Error GoTo Err_Proc


   If SortOrder And LastColName = grdDataGrid.Columns(ColIndex).DataField Then
      SortCol = "[" & grdDataGrid.Columns(ColIndex).DataField & "] desc"
      SortOrder = False
   Else
      SortCol = "[" & grdDataGrid.Columns(ColIndex).DataField & "]"
      SortOrder = True
   End If

   Data1.RecordSource = "Select * From Labels Order By " & SortCol
   Data1.Recordset.Requery
   Data1.Refresh

   LastColName = grdDataGrid.Columns(ColIndex).DataField
   Me.Caption = Data1.RecordSource

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmDataGrid", "grdDataGrid_HeadClick"
   Err.Clear
   Resume Exit_Here

End Sub

Private Sub grdDataGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)

   Static StartBkMark As Variant
   Dim vBkMark        As Variant

   On Local Error Resume Next

   If Button <> vbLeftButton Then Exit Sub

   vBkMark = Data1.Recordset.Bookmark

   If grdDataGrid.SelBookmarks.Count > 0 Then
      If Shift <> vbShiftMask Then
         StartBkMark = Data1.Recordset.Bookmark
      Else
         If IsEmpty(StartBkMark) Then Exit Sub
         Data1.Recordset.Bookmark = StartBkMark
         Do
            grdDataGrid.SelBookmarks.Add Data1.Recordset.Bookmark
            If StartBkMark < vBkMark Then
               Data1.Recordset.MoveNext
               If vBkMark <= Data1.Recordset.Bookmark Then Exit Do
            Else
               Data1.Recordset.MovePrevious
               If vBkMark >= Data1.Recordset.Bookmark Then Exit Do
            End If
         Loop
         Set StartBkMark = Nothing
      End If
   End If

   On Local Error GoTo 0

End Sub

Public Sub PrintGridLabels(ByVal StartCol As Long, ByVal StartRow As Long)

   Dim EndOfFile         As Boolean
   Dim Down              As Long
   Dim HoldPlace         As Variant
   Dim hTabPos           As Single
   Dim VTabPos           As Single
   Dim i                 As Long
   Dim n                 As Long
   Dim vBkMark           As Variant
   Dim cBkMark           As Long
   Dim Line1()           As String
   Dim Line2()           As String
   Dim Line3()           As String
   Dim Line4()           As String
   Dim line5()           As String


   PrintingFlag = True
   PrintingShow True
   DoEvents

   Printer.ScaleMode = vbInches

   On Local Error Resume Next
   Call SetUpPage

   hTabPos = SideMargin
   VTabPos = TopMargin

   HoldPlace = Data1.Recordset.Bookmark

   'cBkMark = grdDataGrid.SelBookmarks.Count
   cBkMark = 0
   vBkMark = grdDataGrid.SelBookmarks(cBkMark)
   Data1.Recordset.Bookmark = vBkMark

   Do
      ReDim Line1(NoAcross) As String
      ReDim Line2(NoAcross) As String
      ReDim Line3(NoAcross) As String
      ReDim Line4(NoAcross) As String
      ReDim line5(NoAcross) As String

      For i = StartCol To NoAcross
         PrintingShow , Data1.Recordset!Line1

         Line1(i) = FixStringCase(Data1.Recordset!Line1 & vbNullString, PrintCaseType)
         Line2(i) = FixStringCase(Data1.Recordset!Line2 & vbNullString, PrintCaseType)
         Line3(i) = FixStringCase(Data1.Recordset!Line3 & vbNullString, PrintCaseType)
         Line4(i) = FixStringCase(Data1.Recordset!Line4 & vbNullString, PrintCaseType)
         If Trim$(Line4(i)) = vbNullString Then
            Line3(i) = Trim$(Line3(i))
            Line3(i) = Left$(Line3(i), Len(Line3(i)) - 1) & UCase$(Right$(Line3(i), 1))
         Else
            Line4(i) = Trim$(Line4(i))
            Line4(i) = Left$(Line4(i), Len(Line4(i)) - 1) & UCase$(Right$(Line4(i), 1))
         End If
         line5(i) = Data1.Recordset!ZipCode

         cBkMark = cBkMark + 1
         If cBkMark = grdDataGrid.SelBookmarks.Count Then
            EndOfFile = True
            Exit For
         End If
         vBkMark = grdDataGrid.SelBookmarks(cBkMark)
         Data1.Recordset.Bookmark = vBkMark

      Next i
      GoSub PrintLines

   Loop Until EndOfFile

   Printer.EndDoc

   PrintingShow False
   Data1.Recordset.Bookmark = HoldPlace

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
      If LenB(Line4(i)) = 0 Then
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
      If Line4(i) > "" Then
         Printer.CurrentX = hTabPos
         n = Len(Line4(i))
         If n > 0 Then
            Line4(i) = Left$(Line4(i), n - 2) & UCase$(Right$(Line4(i), 2))
         End If
         Printer.Print Line4(i); "  "; line5(i);
      End If
      If i = NoAcross Then Printer.Print Else hTabPos = hTabPos + HorzPitch
   Next i

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

