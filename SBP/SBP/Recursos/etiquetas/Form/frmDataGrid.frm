VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDataGrid 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4845
   ClientLeft      =   1305
   ClientTop       =   2340
   ClientWidth     =   8760
   Icon            =   "frmDataGrid.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleMode       =   0  'User
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
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
         TabIndex        =   2
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
         MICON           =   "frmDataGrid.frx":000C
         PICN            =   "frmDataGrid.frx":0028
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
         TabIndex        =   3
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
         MICON           =   "frmDataGrid.frx":0456
         PICN            =   "frmDataGrid.frx":0472
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
         TabIndex        =   4
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
         MICON           =   "frmDataGrid.frx":0A14
         PICN            =   "frmDataGrid.frx":0A30
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
         TabIndex        =   5
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
         MICON           =   "frmDataGrid.frx":0EF6
         PICN            =   "frmDataGrid.frx":0F12
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
         TabIndex        =   6
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
         MICON           =   "frmDataGrid.frx":13D8
         PICN            =   "frmDataGrid.frx":13F4
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
         TabIndex        =   7
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
         MICON           =   "frmDataGrid.frx":19C2
         PICN            =   "frmDataGrid.frx":19DE
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
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Align           =   1  'Align Top
      Bindings        =   "frmDataGrid.frx":1EA4
      Height          =   2025
      Left            =   0
      TabIndex        =   1
      Top             =   630
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   3572
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
      RowDividerStyle =   6
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "LINE1"
         Caption         =   "LINE1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "LINE2"
         Caption         =   "LINE2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "LINE3"
         Caption         =   "LINE3"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "LINE4"
         Caption         =   "LINE4"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "ZIPCODE"
         Caption         =   "ZIPCODE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            WrapText        =   -1  'True
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column02 
            WrapText        =   -1  'True
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column03 
            WrapText        =   -1  'True
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column04 
            WrapText        =   -1  'True
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1500.095
         EndProperty
      EndProperty
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
  Dim i As Integer, c As Integer
  Dim vBkMark As Variant

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
  Dim sFilterStr As String
  Dim FieldName As String
  Dim FilterType As Integer
  Dim i As Byte

    On Error GoTo FilterErr
    
    With frmFilterOptions
        .Move cmdFilter.Left + 100, cmdFilter.Top + cmdFilter.Height + 350
        .Show vbModal
        sFilterStr = Trim(.Text1)
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
    MsgBox "Error:" & Err & " " & Err.Description
End Sub
Public Sub PrintGridLabels(ByVal StartCol As Integer, ByVal StartRow As Integer)
 Dim EndOfFile As Boolean
 Dim Down As Integer, HoldPlace As Variant
 Dim hTabPos As Single, VTabPos As Single
 Dim i As Integer, n As Integer
 Dim vBkMark As Variant
 Dim cBkMark As Long
 Dim Line1() As String
 Dim Line2() As String
 Dim Line3() As String
 Dim Line4() As String
 Dim line5() As String
    
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
           If Trim(Line4(i)) = vbNullString Then
               Line3(i) = Trim(Line3(i))
               Line3(i) = Left(Line3(i), Len(Line3(i)) - 1) & UCase(Right(Line3(i), 1))
           Else
               Line4(i) = Trim(Line4(i))
               Line4(i) = Left(Line4(i), Len(Line4(i)) - 1) & UCase(Right(Line4(i), 1))
           End If
           line5(i) = Data1.Recordset!ZipCode
           
           cBkMark = cBkMark + 1
           If cBkMark = grdDataGrid.SelBookmarks.Count Then EndOfFile = True: Exit For
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
    For i% = StartCol To NoAcross
        Printer.CurrentX = hTabPos
        Printer.Print Line1(i%);
        If i% = NoAcross Then Printer.Print Else hTabPos = hTabPos + HorzPitch
    Next i%
    
    DoEvents
    hTabPos = SideMargin
    If StartCol > 1 Then hTabPos = (hTabPos + HorzPitch) * (StartCol - 1)
    For i% = StartCol To NoAcross
        Printer.CurrentX = hTabPos
        Printer.Print Line2(i%);
        If i% = NoAcross Then Printer.Print Else hTabPos = hTabPos + HorzPitch
    Next i%
    
    DoEvents
    hTabPos = SideMargin
    If StartCol > 1 Then hTabPos = (hTabPos + HorzPitch) * (StartCol - 1)
    For i% = StartCol To NoAcross
        Printer.CurrentX = hTabPos
        If Line4(i%) = "" Then
            n% = Len(Line3(i%))
            If n > 0 Then
                Line3(i%) = Left(Line3(i%), n% - 2) & UCase(Right(Line3(i), 2))
            End If
            Printer.Print Line3(i%); "  " & line5(i%);
        Else
            Printer.Print Line3(i%);
        End If
        If i% = NoAcross Then Printer.Print Else hTabPos = hTabPos + HorzPitch
    Next i%
    
    DoEvents
    hTabPos = SideMargin
    If StartCol > 1 Then hTabPos = (hTabPos + HorzPitch) * (StartCol - 1)
    For i% = StartCol To NoAcross
        If Line4(i%) > "" Then
            Printer.CurrentX = hTabPos
            n% = Len(Line4(i%))
            If n > 0 Then
                Line4(i%) = Left(Line4(i%), n% - 2) & UCase(Right(Line4(i), 2))
            End If
            Printer.Print Line4(i%); "  "; line5(i%);
        End If
        If i% = NoAcross Then Printer.Print Else hTabPos = hTabPos + HorzPitch
    Next i%
    
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


Private Sub cmdPaste_Click()
  Dim i As Integer
  Dim vBkMark As Variant
  Dim SQLstmt As String
  Dim tLine1 As String, tLine2 As String
  Dim tLine3 As String, tLine4 As String
  Dim tZipCode As String

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
    If grdDataGrid.SelBookmarks.Count > 0 Then
        With frmPrinterSetUp
            .FraChoices.Visible = False
            .ReportType = 7
            .Show vbModal
        End With
    End If
    QuitCommand = False
End Sub

Private Sub cmdRefresh_Click()
    On Error GoTo RefErr
    
    Data1.RecordSource = "select * from labels order by [zipcode]"
    Data1.Refresh
    
    Exit Sub
    
RefErr:
    MsgBox "Error:" & Err & " " & Err.Description
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
    MsgBox "Error:" & Err & " " & Err.Description
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call OpenDB(ActiveDB)
    Call OpenRS(ActiveRS, Data1.RecordSource, ActiveDB)
    ADOFindFirst ActiveRS, "[ID]=" & Format(Data1.Recordset!ID)
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

Private Sub grdDataGrid_BeforeUpdate(Cancel As Integer)
    'If MsgBox("Commit changes?", vbYesNo + vbQuestion) <> vbYes Then
    '    Cancel = True
    'End If
End Sub

Private Sub grdDataGrid_DblClick()
    If PasteFileOpen Then
        cmdPaste_Click
    Else
        cmdClose_Click
    End If
End Sub


Private Sub grdDataGrid_HeadClick(ByVal ColIndex As Integer)
  Dim SortCol As String
  Static SortOrder As Boolean, LastColName As String

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
    
End Sub

Private Sub grdDataGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim vBkMark As Variant
  Static StartBkMark As Variant

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
