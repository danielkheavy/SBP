Attribute VB_Name = "Module20"
'*
'* Author: E. J. Bantz Jr.
'* Copyright: None, use and distribute freely ...
'* E-Mail: ejbantz@usa.net
'* Web: http://www.inlink.com/~ejbantz

'// ------------------------------------------------------------------
'//  Windows API Constants / Types / Declarations
'// ------------------------------------------------------------------
Public Const WS_BORDER = &H800000

Public Const WS_CAPTION = &HC00000

Public Const WS_SYSMENU = &H80000

Public Const WS_CHILD = &H40000000

Public Const WS_VISIBLE = &H10000000

Public Const WS_OVERLAPPED = &H0&

Public Const WS_MINIMIZEBOX = &H20000

Public Const WS_MAXIMIZEBOX = &H10000

Public Const WS_THICKFRAME = &H40000

Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)

Public Const SWP_NOMOVE = &H2

Public Const SWP_NOSIZE = 1

Public Const SWP_NOZORDER = &H4

Public Const HWND_BOTTOM = 1

Public Const HWND_TOPMOST = -1

Public Const HWND_NOTOPMOST = -2

Public Const SM_CYCAPTION = 4

Public Const SM_CXFRAME = 32

Public Const SM_CYFRAME = 33

Public Const WS_EX_TRANSPARENT = &H20&

Public Const GWL_STYLE = (-16)

Declare Function SetWindowLong _
        Lib "user32" _
        Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                ByVal nIndex As Long, _
                                ByVal dwNewLong As Long) As Long

'// Memory manipulation
Declare Function lStrCpy _
        Lib "kernel32" _
        Alias "lstrcpyA" (ByVal lpString1 As Long, _
                          ByVal lpString2 As Long) As Long
Declare Function lStrCpyn _
        Lib "kernel32" _
        Alias "lstrcpynA" (ByVal lpString1 As Any, _
                           ByVal lpString2 As Long, _
                           ByVal iMaxLength As Long) As Long
Declare Sub RtlMoveMemory _
        Lib "kernel32" (ByVal hpvDest As Long, _
                        ByVal hpvSource As Long, _
                        ByVal cbCopy As Long)
Declare Sub hmemcpy _
        Lib "kernel32" (hpvDest As Any, _
                        hpvSource As Any, _
                        ByVal cbCopy As Long)
    
'// Window manipulation
Declare Function SetWindowPos _
        Lib "user32" (ByVal hwnd As Long, _
                      ByVal hWndInsertAfter As Long, _
                      ByVal X As Long, _
                      ByVal Y As Long, _
                      ByVal cX As Long, _
                      ByVal cY As Long, _
                      ByVal wFlags As Long) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hndw As Long) As Boolean
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Declare Function SetWindowText _
        Lib "user32" _
        Alias "SetWindowTextA" (ByVal hwnd As Long, _
                                ByVal lpString As String) As Long

Public lwndC As Long       ' Handle to the Capture Windows

Function MyFrameCallback(ByVal lwnd As Long, ByVal lpVHdr As Long) As Long

    Debug.Print "FrameCallBack"
    
    Dim VideoHeader As VIDEOHDR

    Dim VideoData() As Byte
    
    '//Fill VideoHeader with data at lpVHdr
    RtlMoveMemory VarPtr(VideoHeader), lpVHdr, Len(VideoHeader)
    
    '// Make room for data
    ReDim VideoData(VideoHeader.dwBytesUsed)
    
    '//Copy data into the array
    RtlMoveMemory VarPtr(VideoData(0)), VideoHeader.lpData, VideoHeader.dwBytesUsed

    Debug.Print VideoHeader.dwBytesUsed
    Debug.Print VideoData
    
End Function

Function MyYieldCallback(lwnd As Long) As Long

    Debug.Print "Yield"

End Function

Function MyErrorCallback(ByVal lwnd As Long, _
                         ByVal iID As Long, _
                         ByVal ipstrStatusText As Long) As Long
    
    If iID = 0 Then Exit Function
    
    Dim sStatusText  As String

    Dim usStatusText As String
    
    'Convert the Pointer to a real VB String
    sStatusText = String$(255, 0)                                      '// Make room for message
    lStrCpy StrPtr(sStatusText), ipstrStatusText                       '// Copy message into String
    sStatusText = Left$(sStatusText, InStr(sStatusText, Chr$(0)) - 1)  '// Only look at left of null
    usStatusText = StrConv(sStatusText, vbUnicode)                     '// Convert Unicode
            
    LogError usStatusText, iID

End Function

Function MyStatusCallback(ByVal lwnd As Long, _
                          ByVal iID As Long, _
                          ByVal ipstrStatusText As Long) As Long

    If iID = 0 Then Exit Function
   
    Dim sStatusText  As String

    Dim usStatusText As String
    
    '// Convert the Pointer to a real VB String
    sStatusText = String$(255, 0)                                      '// Make room for message
    lStrCpy StrPtr(sStatusText), ipstrStatusText                       '// Copy message into String
    sStatusText = Left$(sStatusText, InStr(sStatusText, Chr$(0)) - 1)  '// Only look at left of null
    usStatusText = StrConv(sStatusText, vbUnicode)                     '// Convert Unicode
    
    'frmMain.StatusBar.SimpleText = usStatusText
    Debug.Print "Status: ", usStatusText, iID

    Select Case iID '
    
    End Select

End Function

Sub ResizeCaptureWindow(ByVal lwnd As Long)

    Dim CAPSTATUS      As CAPSTATUS

    Dim lCaptionHeight As Long

    Dim lX_Border      As Long

    Dim lY_Border      As Long
    
    lCaptionHeight = GetSystemMetrics(SM_CYCAPTION)
    lX_Border = GetSystemMetrics(SM_CXFRAME)
    lY_Border = GetSystemMetrics(SM_CYFRAME)
    
    '// Get the capture window attributes .. width and height
    If capGetStatus(lwnd, VarPtr(CAPSTATUS), Len(CAPSTATUS)) Then
        
        '// Resize the capture window to the capture sizes
        SetWindowPos lwnd, HWND_BOTTOM, 0, 0, CAPSTATUS.uiImageWidth + (lX_Border * 2), CAPSTATUS.uiImageHeight + lCaptionHeight + (lY_Border * 2), SWP_NOMOVE Or SWP_NOZORDER

    End If

    Debug.Print "Resize Window."

End Sub

Function MyVideoStreamCallback(lwnd As Long, lpVHdr As Long) As Long

    Beep  '// Replace this with your code!
  
End Function

Function MyWaveStreamCallback(lwnd As Long, lpVHdr As Long) As Long

    Debug.Print "WaveStream"

End Function

Sub LogError(txtError As String, lID As Long)

    'frmMain.StatusBar.SimpleText = txtError
    Debug.Print "Error: ", txtError, lID
 
End Sub

