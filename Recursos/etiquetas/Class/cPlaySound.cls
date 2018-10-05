VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlaySound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/*************************************/
'/* Author: Morgan Haueisen
'/* Copyright (c) 1997-2002
'/*************************************/

Option Explicit

'/*  flag values for uFlags parameter
Private Const SND_SYNC = &H0              '/*  play synchronously (default)
Private Const SND_ASYNC = &H1             '/*  play asynchronously
Private Const SND_NODEFAULT = &H2         '/*  silence not default, if sound not found
Private Const SND_MEMORY = &H4            '/*  lpszSoundName points to a memory file
Private Const SND_ALIAS = &H10000         '/*  name is a WIN.INI [sounds] entry
Private Const SND_FILENAME = &H20000      '/*  name is a file name
Private Const SND_RESOURCE = &H40004      '/*  name is a resource name or atom
Private Const SND_ALIAS_ID = &H110000     '/*  name is a WIN.INI [sounds] entry identifier
Private Const SND_ALIAS_START = 0         '/*  must be > 4096 to keep strings in same section of resource file
Private Const SND_LOOP = &H8              '/*  loop the sound until next sndPlaySound
Private Const SND_NOSTOP = &H10           '/*  don't stop any currently playing sound
Private Const SND_VALID = &H1F            '/*  valid flags          / ;Internal /
Private Const SND_NOWAIT = &H2000         '/*  don't wait if the driver is busy
Private Const SND_VALIDFLAGS = &H17201F   '/*  Set of valid flag bits.  Anything outside this range will raise an error
Private Const SND_RESERVED = &HFF000000   '/*  In particular these flags are reserved
Private Const SND_TYPE_MASK = &H170007

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function PlaySoundData Lib "winmm.dll" Alias "PlaySoundA" (lpData As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function waveOutGetNumDevs& Lib "winmm.dll" ()

'/*  waveform audio error return values
Private Const WAVERR_BASE = 32
Private Const WAVERR_BADFORMAT = (WAVERR_BASE + 0)       '/*  unsupported wave format
Private Const WAVERR_STILLPLAYING = (WAVERR_BASE + 1)    '/*  still something playing
Private Const WAVERR_UNPREPARED = (WAVERR_BASE + 2)      '/*  header not prepared
Private Const WAVERR_SYNC = (WAVERR_BASE + 3)            '/*  device is synchronous
Private Const WAVERR_LASTERROR = (WAVERR_BASE + 3)       '/*  last error in range

Private m_snd() As Byte

Public Function DevicesPresent() As Boolean

    On Error GoTo Err_Proc

  Dim K, NumDevs, msg$, Rv&, V&

    '/* Query The System As To Whether It Can Play Wave's.
    '/* use The "waveOutGetNumDevs" API Function.
    NumDevs = waveOutGetNumDevs()

    If NumDevs = 0 Then
       DevicesPresent = False
    Else
       DevicesPresent = True
    End If


Exit_Here:
    Exit Function

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "clsPlaySound", "DevicesPresent"
    Err.Clear
    Resume Exit_Here

End Function

Public Function PlaySoundResource(ByVal SndID As Long, Optional SndType As String = "SOUND", Optional Async As Boolean = False) As Long
 Dim Flags As Long
    
    On Local Error Resume Next
    If DevicesPresent Then
        If Async Then
            Flags = SND_MEMORY Or SND_ASYNC Or SND_NODEFAULT
        Else
            Flags = SND_MEMORY Or SND_SYNC Or SND_NODEFAULT
        End If
        SndType = UCase(SndType)
        m_snd = LoadResData(SndID, SndType)
        PlaySoundResource = PlaySoundData(m_snd(0), 0, Flags)
    Else
        Beep
    End If
    On Local Error GoTo 0
    
End Function

Public Sub PlaySoundFile(ByVal PathWAVname As String, Optional Async As Boolean = True, Optional sLoop As Boolean = False)
  Dim rValue As Long, Flags As Long
    
    On Local Error Resume Next
    If DevicesPresent Then
        If Async Then Flags = SND_ASYNC
        If sLoop Then Flags = Flags Or SND_LOOP
    
        If Dir(PathWAVname) > vbNullString Then
            rValue = PlaySound(CStr(PathWAVname), App.hInstance, Flags)
            If rValue = False Then Beep
        Else
            Beep
        End If
    Else
        Beep
    End If
    On Local Error GoTo 0
    
End Sub

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
