VERSION 5.00
Begin VB.Form frmain 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control Camara"
   ClientHeight    =   8475
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   13395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   13395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu kiarc 
      Caption         =   "&Archivo"
      Begin VB.Menu fdk9922 
         Caption         =   "&Asignar"
      End
   End
   Begin VB.Menu kll991 
      Caption         =   "&Editar"
      Begin VB.Menu jm9933 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu fdlo8912 
      Caption         =   "&Control"
      Begin VB.Menu dfj88sta 
         Caption         =   "&Grabar"
      End
      Begin VB.Menu mnuDisplay 
         Caption         =   "&Display"
      End
      Begin VB.Menu mnuFormat 
         Caption         =   "&Formato"
      End
      Begin VB.Menu mnuSource 
         Caption         =   "&Fuente"
      End
      Begin VB.Menu mnuCompression 
         Caption         =   "&Compresion"
      End
      Begin VB.Menu fkol99 
         Caption         =   "&Selecciona Driver"
      End
      Begin VB.Menu mnuScale 
         Caption         =   "&Escala"
      End
   End
   Begin VB.Menu flo442 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*
'* Author: E. J. Bantz Jr.
'* Copyright: None, use and distribute freely ...
'* E-Mail: ej@bantz.com
'* Web: http://ej.bantz.com
'*
Option Explicit

Private Sub dfj88sta_Click()
' /*
'  * If Start is selected from the menu, start Streaming capture.
'  * The streaming capture is terminated when the Escape key is pressed
'  */
    
    Dim sFileName As String
    Dim CAP_PARAMS As CAPTUREPARMS
    
    capCaptureGetSetup lwndC, VarPtr(CAP_PARAMS), Len(CAP_PARAMS)
    
    CAP_PARAMS.dwRequestMicroSecPerFrame = (1 * (10 ^ 6)) / 30  ' 30 Frames per second
    CAP_PARAMS.fMakeUserHitOKToCapture = True
    CAP_PARAMS.fCaptureAudio = False
    
    capCaptureSetSetup lwndC, VarPtr(CAP_PARAMS), Len(CAP_PARAMS)
    
    sFileName = globalpath & "\" & Format(Now, "ddmmyy") & ".avi"
    'MsgBox sFileName
    
    capCaptureSequence lwndC  ' Start Capturing!
    capFileSaveAs lwndC, sFileName  ' Copy video from swap file into a real file.

End Sub

Private Sub fdk9922_Click()

 Dim sFile As String * 250
 Dim lSize As Long
 
 '// Setup swap file for capture
 lSize = 1000000
 sFile = "C:\TEMP.AVI"
 capFileSetCaptureFile lwndC, sFile
 capFileAlloc lwndC, lSize

End Sub

Private Sub fkol99_Click()
    capDlgVideoSource lwndC

End Sub

Private Sub flo442_Click()
frmain.Hide
Unload frmain
End Sub

Private Sub Form_Load()
    
    Dim lpszName As String * 100
    Dim lpszVer As String * 100
    Dim Caps As CAPDRIVERCAPS
        
    'cargar_drivers
    
    '//Create Capture Window
        capGetDriverDescriptionA 0, lpszName, 100, lpszVer, 100  '// Retrieves driver info
    lwndC = capCreateCaptureWindowA(lpszName, WS_CAPTION Or WS_THICKFRAME Or WS_VISIBLE Or WS_CHILD, 0, 0, 160, 120, Me.hWnd, 0)

    
    '// Set title of window to name of driver
    'SetWindowText lwndC, lpszName
    SetWindowText lwndC, "CAMARA EN VIVO"
    
    '// Set the video stream callback function
    capSetCallbackOnStatus lwndC, AddressOf MyStatusCallback
    capSetCallbackOnError lwndC, AddressOf MyErrorCallback
    
    '// Connect the capture window to the driver
    If capDriverConnect(lwndC, 0) Then
        '
        '/////
        '// Only do the following if the connect was successful.
        '// if it fails, the error will be reported in the call
        '// back function.
        '/////
        '// Get the capabilities of the capture driver
        capDriverGetCaps lwndC, VarPtr(Caps), Len(Caps)
        
        '// If the capture driver does not support a dialog, grey it out
        '// in the menu bar.
        If Caps.fHasDlgVideoSource = 0 Then mnuSource.Enabled = False
        If Caps.fHasDlgVideoFormat = 0 Then mnuFormat.Enabled = False
        If Caps.fHasDlgVideoDisplay = 0 Then mnuDisplay.Enabled = False
        
        '// Turn Scale on
        capPreviewScale lwndC, True
            
        '// Set the preview rate in milliseconds
        capPreviewRate lwndC, 66
        
        '// Start previewing the image from the camera
        capPreview lwndC, True
            
        '// Resize the capture window to show the whole image
        ResizeCaptureWindow lwndC

    End If


End Sub

Private Sub Form_Unload(Cancel As Integer)

    '// Disable all callbacks
    capSetCallbackOnError lwndC, vbNull
    capSetCallbackOnStatus lwndC, vbNull
    capSetCallbackOnYield lwndC, vbNull
    capSetCallbackOnFrame lwndC, vbNull
    capSetCallbackOnVideoStream lwndC, vbNull
    capSetCallbackOnWaveStream lwndC, vbNull
    capSetCallbackOnCapControl lwndC, vbNull
    

End Sub

Private Sub mnuAllocate_Click()

 Dim sFile As String * 250
 Dim lSize As Long
 
 '// Setup swap file for capture
 lSize = 1000000
 sFile = "C:\TEMP.AVI"
 capFileSetCaptureFile lwndC, sFile
 capFileAlloc lwndC, lSize
 
End Sub

Private Sub mnuAlwaysVisible_Click()
    
    'mnuAlwaysVisible.Checked = Not (mnuAlwaysVisible.Checked)
    
    'If mnuAlwaysVisible.Checked Then
    '    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    'Else
    '    SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    'End If


End Sub

Private Sub mnuCompression_Click()
'   /*
'   * Display the Compression dialog when "Compression" is selected from
'   * the menu bar.
'   */
    
    capDlgVideoCompression lwndC

End Sub

Private Sub mnuCopy_Click()

    capEditCopy lwndC
        
End Sub

Private Sub jm9933_Click()
    capEditCopy lwndC

End Sub

Private Sub mnuDisplay_Click()
'   /*
'   * Display the Video Display dialog when "Display" is selected from
'   * the menu bar.
'   */

    capDlgVideoDisplay lwndC
    
End Sub

Private Sub mnuExit_Click()

    Unload Me
    
End Sub

Private Sub mnuFormat_Click()
'  /*
'   * Display the Video Format dialog when "Format" is selected from the
'   * menu bar.
'   */

    capDlgVideoFormat lwndC
    ResizeCaptureWindow lwndC

End Sub

Private Sub mnuPreview_Click()

    'frmMain.StatusBar.SimpleText = vbNullString
    'mnuPreview.Checked = Not (mnuPreview.Checked)
    'capPreview lwndC, mnuPreview.Checked
    
End Sub

Private Sub mnuScale_Click()
    
    mnuScale.Checked = Not (mnuScale.Checked)
    capPreviewScale lwndC, mnuScale.Checked
    
    If mnuScale.Checked Then
       SetWindowLong lwndC, GWL_STYLE, WS_THICKFRAME Or WS_CAPTION Or WS_VISIBLE Or WS_CHILD
    Else
       SetWindowLong lwndC, GWL_STYLE, WS_BORDER Or WS_CAPTION Or WS_VISIBLE Or WS_CHILD
    End If

    ResizeCaptureWindow lwndC
    
End Sub

Private Sub mnuSelect_Click()
    
    'frmSelect.Show vbModal, Me

End Sub

Private Sub mnuSource_Click()
    capDlgVideoSource lwndC

End Sub

Private Sub mnuStart_Click()
' /*
'  * If Start is selected from the menu, start Streaming capture.
'  * The streaming capture is terminated when the Escape key is pressed
'  */
    
    Dim sFileName As String
    Dim CAP_PARAMS As CAPTUREPARMS
    
    capCaptureGetSetup lwndC, VarPtr(CAP_PARAMS), Len(CAP_PARAMS)
    
    CAP_PARAMS.dwRequestMicroSecPerFrame = (1 * (10 ^ 6)) / 30  ' 30 Frames per second
    CAP_PARAMS.fMakeUserHitOKToCapture = True
    CAP_PARAMS.fCaptureAudio = False
    
    capCaptureSetSetup lwndC, VarPtr(CAP_PARAMS), Len(CAP_PARAMS)
    sFileName = globalpath & "\" & Format(Now, "ddmmyy") & ".avi"
    'MsgBox sFileName
    'sFileName = "C:\myvideo.avi"
    
    capCaptureSequence lwndC  ' Start Capturing!
    capFileSaveAs lwndC, sFileName  ' Copy video from swap file into a real file.

End Sub


