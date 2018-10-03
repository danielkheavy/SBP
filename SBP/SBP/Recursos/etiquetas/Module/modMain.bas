Attribute VB_Name = "modMain"
Option Explicit

Public AppPath               As String
Public dbPath                As String
Public AddingNew             As Boolean
Public AddNumber             As Long
'Public FormDone              As Boolean
Public QuitCommand           As Boolean
'Public CreateNewFile         As Boolean

Public PasteDB               As ADODB.Connection
Public PasteRS               As ADODB.Recordset
Public ActiveDB              As ADODB.Connection
Public ActiveRS              As ADODB.Recordset

Public PasteFile             As Boolean
Public PasteFileOpen         As Boolean
Public PrintCaseType         As Long
Public PrintingFlag          As Boolean

Public cScreen               As New clsScreenSize
Public cValidate             As New clsValidate
Public cPlay                 As New clsPlaySound

Public SideMargin            As Single
Public TopMargin             As Single
Public VerPitch              As Single
Public HorzPitch             As Single
Public NoAcross              As Long
Public NoDown                As Long
Public SchemeID              As Long

Public dFontStyle            As String
Public dFontSize             As Long
Public dFontName             As String
Public dFontUnderline        As Boolean
Public dFontStrikeThru       As Boolean
Public dPrintScheme          As String

Public Function FixStringCase(ByVal tString As String, _
                              Optional cCaseType As Long = vbProperCase) As String

   If cCaseType = vbUpperCase Then

      tString = UCase$(tString)
   End If
   FixStringCase = tString

End Function

Public Sub PrintAllLabels(ByVal PrintAll As Boolean, _
                          ByVal StartCol As Long, _
                          ByVal StartRow As Long)

   Dim Down              As Long
   Dim HoldPlace         As Variant
   Dim hTabPos           As Single
   Dim VTabPos           As Single
   Dim i                 As Long
   Dim n                 As Long
   Dim Line1()           As String
   Dim Line2()           As String
   Dim Line3()           As String
   Dim Line4()           As String
   Dim line5()           As String
   Dim EndOfFile         As Boolean



   PrintingFlag = True
   PrintingShow True
   DoEvents

   Printer.ScaleMode = vbInches

   On Local Error Resume Next
   Call SetUpPage

   hTabPos = SideMargin
   VTabPos = TopMargin

   HoldPlace = ActiveRS.Bookmark
   If PrintAll Then ActiveRS.MoveFirst

   Do

      ReDim Line1(NoAcross) As String
      ReDim Line2(NoAcross) As String
      ReDim Line3(NoAcross) As String
      ReDim Line4(NoAcross) As String
      ReDim line5(NoAcross) As String

      For i = StartCol To NoAcross
         PrintingShow , ActiveRS!Line1

         Line1(i) = FixStringCase(ActiveRS!Line1 & vbNullString, PrintCaseType)
         Line2(i) = FixStringCase(ActiveRS!Line2 & vbNullString, PrintCaseType)
         Line3(i) = FixStringCase(ActiveRS!Line3 & vbNullString, PrintCaseType)
         Line4(i) = FixStringCase(ActiveRS!Line4 & vbNullString, PrintCaseType)
         If Trim$(Line4(i)) = vbNullString Then
            Line3(i) = Trim$(Line3(i))
            Line3(i) = Left$(Line3(i), Len(Line3(i)) - 1) & UCase$(Right$(Line3(i), 1))
         Else
            Line4(i) = Trim$(Line4(i))
            Line4(i) = Left$(Line4(i), Len(Line4(i)) - 1) & UCase$(Right$(Line4(i), 1))
         End If
         line5(i) = ActiveRS!ZipCode
         ActiveRS.MoveNext
         If ActiveRS.EOF Then
            EndOfFile = True
            Exit For
         End If
      Next i
      GoSub PrintLines

   Loop Until EndOfFile

   Printer.EndDoc

   PrintingShow False
   ActiveRS.Bookmark = HoldPlace

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
      If Not PrintAll Then
         EndOfFile = True
      Else
         Call SetUpNewPage
         Down = False
         VTabPos = TopMargin
      End If
   Else
      VTabPos = VTabPos + VerPitch
   End If
   StartCol = 1

   DoEvents
   If PrintingFlag = False Then EndOfFile = True

   Return

End Sub

Public Sub PrintingShow(Optional ByVal dShow As Boolean = True, _
                        Optional ByVal Progress As String = vbNullString)



   On Error GoTo Err_Proc

   If dShow Then
      frmMain.PicPrinting.Visible = True
      frmMain.PicPrinting.ZOrder
      If Progress > vbNullString Then
         frmMain.lblPrintUpdate.Text = Progress
         frmMain.lblPrintUpdate.Refresh
      End If
   Else
      frmMain.PicPrinting.Visible = False
      frmMain.lblPrintUpdate.Text = "Initializing Printer.."
   End If

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "modMain", "PrintingShow"
   Err.Clear
   Resume Exit_Here

End Sub

Public Sub PrintNamesOnly(ByVal PrintAll As Boolean, _
                          ByVal StartCol As Long, _
                          ByVal StartRow As Long)

   Dim Down              As Long
   Dim HoldPlace         As Variant
   Dim hTabPos           As Single
   Dim VTabPos           As Single
   Dim i                 As Long
   Dim n                 As Long
   Dim Line1()           As String
   Dim EndOfFile         As Boolean



   PrintingFlag = True
   PrintingShow True

   Printer.ScaleMode = vbInches

   On Local Error Resume Next
   Call SetUpPage

   hTabPos = SideMargin
   VTabPos = TopMargin

   HoldPlace = ActiveRS.Bookmark
   If PrintAll Then ActiveRS.MoveFirst

   Do

      ReDim Line1(NoAcross) As String

      For i = StartCol To NoAcross
         PrintingShow , ActiveRS!Line1

         Line1(i) = FixStringCase(ActiveRS!Line1 & vbNullString, PrintCaseType)
         ActiveRS.MoveNext
         If ActiveRS.EOF Then
            EndOfFile = True
            Exit For
         End If
      Next i
      GoSub PrintLines

   Loop Until EndOfFile

   Printer.EndDoc

   PrintingShow False
   ActiveRS.Bookmark = HoldPlace

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
   Printer.Print
   Printer.Print

   hTabPos = SideMargin
   If StartCol > 1 Then hTabPos = (hTabPos + HorzPitch) * (StartCol - 1)
   For i = StartCol To NoAcross
      Printer.CurrentX = hTabPos
      Printer.Print Line1(i);
      If i = NoAcross Then Printer.Print Else hTabPos = hTabPos + HorzPitch
   Next i

   If Down = NoDown Then
      If Not PrintAll Then
         EndOfFile = True
      Else
         Call SetUpNewPage
         Down = False
         VTabPos = TopMargin
      End If
   Else
      VTabPos = VTabPos + VerPitch
   End If
   StartCol = 1

   DoEvents
   If PrintingFlag = False Then EndOfFile = True

   Return

End Sub

Public Sub PrintOneX(ByVal StartCol As Long, _
                     ByVal StartRow As Long, _
                     Optional ByVal xTimes As Long = 1)

   Dim hTabPos           As Single
   Dim VTabPos           As Single
   Dim i                 As Long
   Dim n                 As Long
   Dim X                 As Long
   Dim Line1()           As String
   Dim Line2()           As String
   Dim Line3()           As String
   Dim Line4()           As String
   Dim line5()           As String
   Dim Down              As Long



   PrintingFlag = True
   PrintingShow True
   DoEvents

   Printer.ScaleMode = vbInches

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
         PrintingShow , ActiveRS!Line1

         Line1(i) = FixStringCase(ActiveRS!Line1 & vbNullString, PrintCaseType)
         Line2(i) = FixStringCase(ActiveRS!Line2 & vbNullString, PrintCaseType)
         Line3(i) = FixStringCase(ActiveRS!Line3 & vbNullString, PrintCaseType)
         Line4(i) = FixStringCase(ActiveRS!Line4 & vbNullString, PrintCaseType)
         If Trim$(Line4(i)) = vbNullString Then
            Line3(i) = Trim$(Line3(i))
            Line3(i) = Left$(Line3(i), Len(Line3(i)) - 1) & UCase$(Right$(Line3(i), 1))
         Else
            Line4(i) = Trim$(Line4(i))
            Line4(i) = Left$(Line4(i), Len(Line4(i)) - 1) & UCase$(Right$(Line4(i), 1))
         End If
         line5(i) = ActiveRS!ZipCode
         X = X + 1
         If xTimes = X Then Exit For
      Next i
      GoSub PrintLines
      If xTimes = X Then Exit Do
   Loop

   Printer.EndDoc

   PrintingShow False

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

   Return

End Sub

Public Sub SetUpNewPage()


   On Error GoTo Err_Proc

   With Printer
      .NewPage
      .FontName = dFontName
      .FontSize = dFontSize
      .FontStrikethru = dFontStrikeThru
      .FontUnderline = dFontUnderline
      Select Case dFontStyle
      Case "regular"
         .FontBold = False
         .FontItalic = False
      Case "italic"
         .FontBold = False
         .FontItalic = True
      Case "bold"
         .FontBold = True
         .FontItalic = False
      Case "bolditalic"
         .FontBold = True
         .FontItalic = True
      End Select
   End With

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "modMain", "SetUpNewPage"
   Err.Clear
   Resume Exit_Here

End Sub

Public Sub SetUpPage()

   On Error GoTo Err_Proc

   With Printer
      .Orientation = vbPRORPortrait
      .PrintQuality = vbPRPQHigh
      .ScaleMode = vbInches

      .FontName = dFontName
      .FontSize = dFontSize
      .FontStrikethru = dFontStrikeThru
      .FontUnderline = dFontUnderline
      Select Case dFontStyle
      Case "regular"
         .FontBold = False
         .FontItalic = False
      Case "italic"
         .FontBold = False
         .FontItalic = True
      Case "bold"
         .FontBold = True
         .FontItalic = False
      Case "bolditalic"
         .FontBold = True
         .FontItalic = True
      End Select
   End With

Exit_Here:

Exit Sub


Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "modMain", "SetUpPage"
   Err.Clear
   Resume Exit_Here

End Sub

Public Sub txtGotFocus(ByRef X As TextBox)

   X.SelStart = 0

   X.SelLength = Len(X)

End Sub


