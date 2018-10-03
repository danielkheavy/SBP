VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsValidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/*************************************/
'/* Author: Morgan Haueisen
'/* Copyright (c) 1997-2002
'/*************************************/
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long



Public Function ClipNull(InString As String) As String

    On Error GoTo Err_Proc

  Dim intpos As Long
    If Len(InString) Then
        intpos = InStr(InString, vbNullChar)
        If intpos > 0 Then
            ClipNull = Left(InString, intpos - 1)
        Else
            ClipNull = InString
        End If
    End If

Exit_Here:
    Exit Function

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "clsValidate", "ClipNull"
    Err.Clear
    Resume Exit_Here

End Function

Public Sub mTrimString(sString As String)

    On Error GoTo Err_Proc

    '/**************************************/
    '/* Remove double spaces between words */
    '/**************************************/
    Dim LastC As String, i As Long
    Const SpaceC As String = " "
    
    sString = Trim(sString)
    If Len(sString) < 2 Then Exit Sub
    i = 1
    LastC = Left(sString, i)
    
    Do
       If LastC = SpaceC And Mid(sString, i, 1) = SpaceC Then
           sString = Left(sString, i - 1) & Mid(sString, i + 1)
       Else
           LastC = Mid(sString, i, 1)
           i = i + 1
       End If
    Loop Until i >= Len(sString)
 

Exit_Here:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "clsValidate", "mTrimString"
    Err.Clear
    Resume Exit_Here

End Sub

Public Sub FcaseTextBox(MyCTRL As Object)

    On Error GoTo Err_Proc

    '/**************************************/
    '/* Change to formal case while typing */
    '/* Place in MyCTRL_KeyUp event        */
    '/**************************************/
    Dim X As Long
    Dim tString As String
    
    If MyCTRL.Tag = vbNullString Then
        MyCTRL.Tag = MyCTRL.SelStart
        tString = FcaseStringF(MyCTRL.Text, True)
        MyCTRL.Text = LTrim$(tString)
        MyCTRL.SelStart = MyCTRL.Tag
        MyCTRL.Tag = vbNullString
    End If

Exit_Here:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "clsValidate", "FcaseTextBox"
    Err.Clear
    Resume Exit_Here

End Sub

Public Sub AutoMatch(Cbo As ComboBox, KeyAscii As Integer, Optional AllowNoMatch As Boolean = False)

    On Error GoTo Err_Proc

'/* Place Sub in MyCTRL_KeyPress */
'/* Place 'If KeyCode = 46 Then KeyCode = False' in MyCtrl_KeyDown to disable the Delete Key

 Const CB_ERR = -1
 Const CB_FINDSTRING = &H14C
 Dim sBuffer As String, RetVal As Long, i As Long
  
  
    Select Case KeyAscii
    Case 13
        KeyAscii = 0
        SendKeys "{TAB}"
    Case 8
        i = Cbo.SelStart
        i = i - 1
        If i < 0 Then i = 0
        Cbo.SelStart = i
        Cbo.SelLength = Len(Cbo.Text)
    Case Else
        sBuffer = Left(Cbo.Text, Cbo.SelStart) & Chr(KeyAscii)
        RetVal = SendMessage((Cbo.hwnd), CB_FINDSTRING, -1, ByVal sBuffer)
        If RetVal <> CB_ERR Then
            Cbo.ListIndex = RetVal
            Cbo.Text = Cbo.List(RetVal)
            Cbo.SelStart = Len(sBuffer)
            Cbo.SelLength = Len(Cbo.Text)
            KeyAscii = False
        End If
    End Select
    
    If Not AllowNoMatch Then KeyAscii = False
    

Exit_Here:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "clsValidate", "AutoMatch"
    Err.Clear
    Resume Exit_Here

End Sub


Public Function MaskedText(MyCTRL As TextBox, ByVal MyProfile As String) As Boolean
'/* Place in MyCTRL_KeyUp for imediate check */
'/* Place in MyCTRL_KeyPress for 1 char delay check */

 Const GoodMatchColor = &H80000008
 Const BadMatchColor = vbRed

 On Local Error GoTo BadCheckProfile

 Dim MyCursorPlace As Long
 Dim MyLen As Long
 Dim MyPlace As Long
 Dim MyBuffer As String
 Dim MyText As String
 Dim MyChar As String * 1
 Dim MyProfilePlace As Long

    MyPlace = 1
    MyProfilePlace = 1
    
    '/* If there are more characters than allowed then remove them */
    If Len(MyCTRL.Text) > Len(MyProfile) Then
        MyCTRL.Text = Left(MyCTRL.Text, Len(MyProfile))
        MyCTRL.SelStart = Len(MyCTRL.Text)
        Beep
    End If
    
    MyText = MyCTRL.Text
    MyLen = Len(MyText)
    MyCursorPlace = MyCTRL.SelStart
    
    '/* The parser takes the pattern as the transition map. Starting
    '/* at the beginning of the map, it compares the current character
    '/* with the state of the parser
    Do While MyPlace <= MyLen
        MyChar = Mid(MyText, MyPlace, 1)
        Select Case Mid(MyProfile, MyProfilePlace, 1)
            '/* Check if MyProfile calls for a numeric input */
         Case "#"
            If IsNumeric(MyChar) Then
                '/* Add the character to the buffer
                MyBuffer = MyBuffer & MyChar
                '/* Move to the next character
                MyPlace = MyPlace + 1
                '/* Move to the next valid parser state
                MyProfilePlace = MyProfilePlace + 1
                '/* Make sure we are indicating a valid transition state
                MyCTRL.ForeColor = GoodMatchColor
            Else
                '/* The character does not match the parser's state so
                '/* indicate an invalid state and exit the parser
                MyCTRL.ForeColor = BadMatchColor
                GoTo ExitCheckProfile
            End If
         Case "@"
            If Not IsNumeric(MyChar) Then
                '/* Add the character to the buffer
                MyBuffer = MyBuffer & MyChar
                '/* Move to the next character
                MyPlace = MyPlace + 1
                '/* Move to the next valid parser state
                MyProfilePlace = MyProfilePlace + 1
                '/* Make sure we are indicating a valid transition state
                MyCTRL.ForeColor = GoodMatchColor
            Else
                '/* The character does not match the parser's state so
                '/* indicate an invalid state and exit the parser
                MyCTRL.ForeColor = BadMatchColor
                GoTo ExitCheckProfile
            End If
         Case Else
            If MyChar = Mid(MyProfile, MyProfilePlace, 1) Then
                '/* If it Is here Then add the character to the buffer */
                MyBuffer = MyBuffer & MyChar
                '/* Move to next character position */
                MyPlace = MyPlace + 1
                '/* Move to next parser state */
                MyProfilePlace = MyProfilePlace + 1
                '/* Indicate a valid transition state to the user */
                MyCTRL.ForeColor = GoodMatchColor
            Else
                '/* The required character is not present and in this case we
                '/* insert it meeting the requirements of the parser state
                MyBuffer = MyBuffer & Mid(MyProfile, MyProfilePlace, 1)
                '/* Shift the parser to the next state but stay with the current
                '/* character to see if it matches the next state transition
                MyProfilePlace = MyProfilePlace + 1
                '/* Move the input cursor
                '/* over one to account for the added character
                MyCursorPlace = MyCursorPlace + 1
            End If
        End Select
    Loop
    
    '/* If this is a function */
    If Len(MyBuffer) = Len(MyProfile) Then
        MaskedText = True
    Else
        MaskedText = False
    End If
    
    MyCTRL.Text = MyBuffer
    MyCTRL.SelStart = MyCursorPlace
    
    '=======================================
ExitCheckProfile:
    '/* exit code here */
Exit Function


BadCheckProfile:
    MsgBox Err.Description, vbCritical, "CheckProfile"
    '/* If this is a function */
    MaskedText = False
    Resume ExitCheckProfile
End Function
Public Function BooleanToCheck(ByVal InVal As Variant)

    On Error GoTo Err_Proc

 '/* This function takes a boolean value or a zero-length string
 '/* and converts it to the value expected by a VB check box.
    
    If IsNullOrEmpty(InVal) Then
        BooleanToCheck = 2
    Else
        Select Case InVal
        Case ""
         BooleanToCheck = 2
        Case True
         BooleanToCheck = 1
        Case False
         BooleanToCheck = 0
        Case Else
         BooleanToCheck = 0
        End Select
    End If

Exit_Here:
    Exit Function

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "clsValidate", "BooleanToCheck"
    Err.Clear
    Resume Exit_Here

End Function

Function IsNullOrEmpty(ByVal InVal As Variant) As Boolean
'/* This function returns True if the value passed as InVal is
'/* Null or a zero-length string.

    If IsNull(InVal) Then
        IsNullOrEmpty = True
    Else
        IsNullOrEmpty = (Len(InVal) = 0)
    End If
End Function

Function CheckToBoolean(ByVal InVal As Integer)
 '/* This function converts the value of a VB checkbox to a
 '/* boolean value or a zero-length string.

    Select Case InVal
    Case 0
      CheckToBoolean = False
    Case 1
      CheckToBoolean = True
    Case Else
      CheckToBoolean = ""
    End Select
End Function


Public Function MakeNotNull(ByVal InVal As Variant, Optional Numeric As Boolean = False) As Variant
 '/* This function turns a null charecter into a zero-length string
 '/* so the value can be inserted into a text box or other control
 '/* that won't accept null values.

    If Numeric Then
        If IsNull(InVal) Then MakeNotNull = 0 Else MakeNotNull = InVal
    Else
        If IsNull(InVal) Then MakeNotNull = vbNullString Else MakeNotNull = InVal
    End If
End Function
Public Function Max(ByVal a As Variant, ByVal B As Variant)
    If a > B Then Max = a Else Max = B
End Function
Public Function Min(ByVal a As Variant, ByVal B As Variant)
    If a < B Then Min = a Else Min = B
End Function

Public Sub FcaseString(tString As String, Optional NoRTrim As Boolean = False)
    Dim X As Long
    '/*************************************/
    '/* Convert any string to Formal Case */
    '/*************************************/
    On Local Error Resume Next
    If NoRTrim Then
        tString = LTrim(tString)
    Else
        tString = Trim(tString)
    End If
    tString = UCase$(Left$(tString, 1)) & LCase$(Mid$(tString, 2))
    
    '/* Look for space */
    X = InStr(tString, " ")
    If X Then
       Do
           Mid$(tString, X + 1, 1) = UCase$(Mid$(tString, X + 1, 1))
           X = X + 1
           X = InStr(X, tString, " ")
           If X = 0 Or X + 1 > Len(tString) Then Exit Do
       Loop
    End If
    '/* Look for . */
    X = InStr(tString, ".")
    If X Then
       Do
           Mid$(tString, X + 1, 1) = UCase$(Mid$(tString, X + 1, 1))
           X = X + 1
           X = InStr(X, tString, ".")
           If X = 0 Or X + 1 > Len(tString) Then Exit Do
       Loop
    End If
    '/* Look for / */
    X = InStr(tString, "/")
    If X Then
       Do
           Mid$(tString, X + 1, 1) = UCase$(Mid$(tString, X + 1, 1))
           X = X + 1
           X = InStr(X, tString, "/")
           If X = 0 Or X + 1 > Len(tString) Then Exit Do
       Loop
    End If
    '/* Look for & */
    X = InStr(tString, "&")
    If X Then
       Do
           Mid$(tString, X + 1, 1) = UCase$(Mid$(tString, X + 1, 1))
           X = X + 1
           X = InStr(X, tString, "&")
           If X = 0 Or X + 1 > Len(tString) Then Exit Do
       Loop
    End If
    '/* Look for Mc */
    X = InStr(tString, "Mc")
    If X Then
       Do
           Mid$(tString, X + 2, 1) = UCase$(Mid$(tString, X + 2, 1))
           X = X + 2
           X = InStr(X, tString, "Mc")
           If X = 0 Or X + 2 > Len(tString) Then Exit Do
       Loop
    End If
    '/* Look for O' */
    X = InStr(tString, "O'")
    If X Then
       Do
           Mid$(tString, X + 2, 1) = UCase$(Mid$(tString, X + 2, 1))
           X = X + 2
           X = InStr(X, tString, "O'")
           If X = 0 Or X + 2 > Len(tString) Then Exit Do
       Loop
    End If
    '/* Look for - */
    X = InStr(tString, "-")
    If X Then
       Do
           Mid$(tString, X + 1, 1) = UCase$(Mid$(tString, X + 1, 1))
           X = X + 1
           X = InStr(X, tString, "-")
           If X = 0 Or X + 1 > Len(tString) Then Exit Do
       Loop
    End If

End Sub

Public Function FcaseStringF(ByVal tString As String, Optional NoRTrim As Boolean = False, Optional CaseType As Integer = vbProperCase) As String
    Dim X As Long
    '/*************************************/
    '/* Convert any string to Formal Case */
    '/*************************************/
    On Local Error Resume Next
    tString = tString & vbNullString
    If NoRTrim Then
        tString = LTrim(tString)
    Else
        tString = Trim(tString)
    End If
    If tString = vbNullString Then
        FcaseStringF = vbNullString
        Exit Function
    End If
    
    Select Case CaseType
    Case vbUpperCase
        FcaseStringF = UCase(tString)
    Case vbLowerCase
        FcaseStringF = LCase(tString)
    Case vbProperCase
    
        tString = UCase$(Left$(tString, 1)) & LCase$(Mid$(tString, 2))
        '/* Look for space */
        X = InStr(tString, " ")
        If X Then
           Do
               Mid$(tString, X + 1, 1) = UCase$(Mid$(tString, X + 1, 1))
               X = X + 1
               X = InStr(X, tString, " ")
               If X = 0 Or X + 1 > Len(tString) Then Exit Do
           Loop
        End If
        '/* Look for . */
        X = InStr(tString, ".")
        If X Then
           Do
               Mid$(tString, X + 1, 1) = UCase$(Mid$(tString, X + 1, 1))
               X = X + 1
               X = InStr(X, tString, ".")
               If X = 0 Or X + 1 > Len(tString) Then Exit Do
           Loop
        End If
        '/* Look for / */
        X = InStr(tString, "/")
        If X Then
           Do
               Mid$(tString, X + 1, 1) = UCase$(Mid$(tString, X + 1, 1))
               X = X + 1
               X = InStr(X, tString, "/")
               If X = 0 Or X + 1 > Len(tString) Then Exit Do
           Loop
        End If
        '/* Look for & */
        X = InStr(tString, "&")
        If X Then
           Do
               Mid$(tString, X + 1, 1) = UCase$(Mid$(tString, X + 1, 1))
               X = X + 1
               X = InStr(X, tString, "&")
               If X = 0 Or X + 1 > Len(tString) Then Exit Do
           Loop
        End If
        '/* Look for Mc */
        X = InStr(tString, "Mc")
        If X Then
           Do
               Mid$(tString, X + 2, 1) = UCase$(Mid$(tString, X + 2, 1))
               X = X + 2
               X = InStr(X, tString, "Mc")
               If X = 0 Or X + 2 > Len(tString) Then Exit Do
           Loop
        End If
        '/* Look for O' */
        X = InStr(tString, "O'")
        If X Then
           Do
               Mid$(tString, X + 2, 1) = UCase$(Mid$(tString, X + 2, 1))
               X = X + 2
               X = InStr(X, tString, "O'")
               If X = 0 Or X + 2 > Len(tString) Then Exit Do
           Loop
        End If
        '/* Look for - */
        X = InStr(tString, "-")
        If X Then
           Do
               Mid$(tString, X + 1, 1) = UCase$(Mid$(tString, X + 1, 1))
               X = X + 1
               X = InStr(X, tString, "-")
               If X = 0 Or X + 1 > Len(tString) Then Exit Do
           Loop
        End If
    End Select
    FcaseStringF = tString
    
End Function


Public Sub mTrimTextBox(MyCTRL As TextBox)
    '/**************************************/
    '/* Remove double spaces between words */
    '/**************************************/
    Dim X As Long
    Dim tString As String
    
    tString = MyCTRL.Text
    Call mTrimString(tString)
    MyCTRL.Text = tString

End Sub

Public Function IsNumber(ByVal CheckString As String, Optional KeyAscii As Integer = 0, Optional AllowDecPoint As Boolean = True, Optional AllowNegative As Boolean = False) As Boolean

    On Error GoTo Err_Proc

    If KeyAscii > 0 And KeyAscii <> 8 Then
        If Not AllowNegative And KeyAscii = 45 Then KeyAscii = 0
        If Not AllowDecPoint And KeyAscii = 46 Then KeyAscii = 0
        If Not IsNumeric(CheckString & Chr(KeyAscii)) Then
            KeyAscii = False
            IsNumber = False
        Else
            IsNumber = True
        End If
    Else
        IsNumber = IsNumeric(CheckString)
    End If

Exit_Here:
    Exit Function

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "clsValidate", "IsNumber"
    Err.Clear
    Resume Exit_Here

End Function

Private Sub Err_Handler( _
    Optional ByVal DisplayError As Boolean = True, _
    Optional ByVal ErrNumber As String = vbNullString, _
    Optional ByVal ErrDescription As String = vbNullString, _
    Optional ByVal ModuleName As String = vbNullString, _
    Optional ByVal ProcName As String = vbNullString)

  Dim tString As String

    '/* Purpose: Error handling - On Error

    '/* Show Error Message
    If DisplayError Then
        tString = "Error occured: "
        If Len(ErrNumber) > 0 Then tString = tString & ErrNumber & vbNewLine Else tString = tString & vbNewLine
        If Len(ErrDescription) > 0 Then tString = tString & "Description: " & ErrDescription & vbNewLine
        If Len(ModuleName) > 0 Then tString = tString & "Module: " & ModuleName & vbNewLine
        If Len(ProcName) > 0 Then tString = tString & "Function: " & ProcName
        MsgBox tString, vbCritical, App.Title & " - ERROR"
    End If

    '/* Write error log
    Dim fnum As Long
    fnum = FreeFile
    Open App.Path & "\ErrorLog.txt" For Append As #fnum
    Write #fnum, Now, ErrNumber, ErrDescription, ModuleName, ProcName, Environ("username"), Environ("computername")
    Close #fnum
End Sub
