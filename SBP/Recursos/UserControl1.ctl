VERSION 5.00
Begin VB.UserControl EC_Button 
   AutoRedraw      =   -1  'True
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1515
   ScaleHeight     =   63
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   101
End
Attribute VB_Name = "EC_Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Enum eStyle

    XP_Button
    Gradient_Button

End Enum

Enum isbAlign

    [isbCenter] = &H0
    [isbleft] = &H1
    [isbRight] = &H2
    [isbTop] = &H3
    [isbbottom] = &H4

End Enum

Private Type POINT

    X As Long
    Y As Long

End Type

Private Type RECT

    Left As Long
    Top As Long
    Right As Long
    bottom As Long

End Type

Private Type ICONINFO

    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long

End Type

Private Type BITMAP

    bmType       As Long    'LONG   // Specifies the bitmap type. This member must be zero.
    bmWidth      As Long    'LONG   // Specifies the width, in pixels, of the bitmap. The width must
    '   be greater than zero.
    bmHeight     As Long    'LONG   // Specifies the height, in pixels, of the bitmap. The height
    '   must be greater than zero.
    bmWidthBytes As Long    'LONG   // Specifies the number of bytes in each scan line. This value
    '   must be divisible by 2, because Windows assumes that the bit values of a bitmap form an array
    '   that is word aligned.
    bmPlanes     As Integer 'WORD   // Specifies the count of color planes.
    bmBitsPixel  As Integer 'WORD   // Specifies the number of bits required to indicate the color of
    '   a pixel.
    bmBits       As Long    'LPVOID // Points to the location of the bit values for the bitmap. The

    '   bmBits member must be a long pointer to an array of character (1-byte) values.
End Type

Private Type RGBQUAD

    rgbBlue                     As Byte
    rgbGreen                    As Byte
    rgbRed                      As Byte
    rgbReserved                 As Byte

End Type

Private Type BITMAPINFOHEADER

    biSize                      As Long
    biWidth                     As Long
    biHeight                    As Long
    biPlanes                    As Integer
    biBitCount                  As Integer
    biCompression               As Long
    biSizeImage                 As Long
    biXPelsPerMeter             As Long
    biYPelsPerMeter             As Long
    biClrUsed                   As Long
    biClrImportant              As Long

End Type

Private Type BITMAPINFO

    bmiHeader                   As BITMAPINFOHEADER
    bmiColors                   As RGBQUAD

End Type

Private Declare Function DrawState _
                Lib "user32" _
                Alias "DrawStateA" (ByVal hDC As Long, _
                                    ByVal hBrush As Long, _
                                    ByVal lpDrawStateProc As Long, _
                                    ByVal lParam As Long, _
                                    ByVal wParam As Long, _
                                    ByVal X As Long, _
                                    ByVal Y As Long, _
                                    ByVal cX As Long, _
                                    ByVal cY As Long, _
                                    ByVal fuFlags As Long) As Long

Private Declare Function DrawStateText _
                Lib "user32" _
                Alias "DrawStateA" (ByVal hDC As Long, _
                                    ByVal hBrush As Long, _
                                    ByVal lpDrawStateProc As Long, _
                                    ByVal lParam As String, _
                                    ByVal wParam As Long, _
                                    ByVal n1 As Long, _
                                    ByVal n2 As Long, _
                                    ByVal n3 As Long, _
                                    ByVal n4 As Long, _
                                    ByVal un As Long) As Long

Private Declare Function BitBlt _
                Lib "gdi32" (ByVal hDestDC As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long, _
                             ByVal nWidth As Long, _
                             ByVal nHeight As Long, _
                             ByVal hSrcDC As Long, _
                             ByVal XSrc As Long, _
                             ByVal YSrc As Long, _
                             ByVal dwRop As Long) As Long

Private Declare Function SetParent _
                Lib "user32.dll" (ByVal hWndChild As Long, _
                                  ByVal hWndNewParent As Long) As Long

Private Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long

Private Declare Function ReleaseDC _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal hDC As Long) As Long

Private Declare Function DrawIconEx _
                Lib "user32" (ByVal hDC As Long, _
                              ByVal xLeft As Long, _
                              ByVal yTop As Long, _
                              ByVal hIcon As Long, _
                              ByVal cxWidth As Long, _
                              ByVal cyWidth As Long, _
                              ByVal istepIfAniCur As Long, _
                              ByVal hbrFlickerFreeDraw As Long, _
                              ByVal diFlags As Long) As Long

Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long

Private Declare Function GetIconInfo _
                Lib "user32.dll" (ByVal hIcon As Long, _
                                  ByRef piconinfo As ICONINFO) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function CreateCompatibleBitmap _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal nWidth As Long, _
                             ByVal nHeight As Long) As Long

Private Declare Function SelectObject _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal hObject As Long) As Long

Private Declare Function SetPixelV _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long, _
                             ByVal crColor As Long) As Long

Private Declare Function GetPixel _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long) As Long

Private Declare Function DrawEdge _
                Lib "user32" (ByVal hDC As Long, _
                              qrc As RECT, _
                              ByVal edge As Long, _
                              ByVal grfFlags As Long) As Long

Private Declare Function SetRect _
                Lib "user32" (lpRect As RECT, _
                              ByVal X1 As Long, _
                              ByVal Y1 As Long, _
                              ByVal X2 As Long, _
                              ByVal Y2 As Long) As Long

Private Declare Function CopyRect _
                Lib "user32" (lpDestRect As RECT, _
                              lpSourceRect As RECT) As Long

Private Declare Function DrawText _
                Lib "user32" _
                Alias "DrawTextA" (ByVal hDC As Long, _
                                   ByVal lpStr As String, _
                                   ByVal ncount As Long, _
                                   lpRect As RECT, _
                                   ByVal wFormat As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function CreateRectRgn _
                Lib "gdi32" (ByVal X1 As Long, _
                             ByVal Y1 As Long, _
                             ByVal X2 As Long, _
                             ByVal Y2 As Long) As Long

Private Declare Function SetWindowRgn _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal hRgn As Long, _
                              ByVal bRedraw As Boolean) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function OleTranslateColor _
                Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, _
                                    ByVal HPALETTE As Long, _
                                    pccolorref As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetObjectAPI _
                Lib "gdi32" _
                Alias "GetObjectA" (ByVal hObject As Long, _
                                    ByVal ncount As Long, _
                                    lpObject As Any) As Long

Private Declare Function GetDIBits _
                Lib "gdi32" (ByVal aHDC As Long, _
                             ByVal hBitmap As Long, _
                             ByVal nStartScan As Long, _
                             ByVal nNumScans As Long, _
                             lpBits As Any, _
                             lpbi As BITMAPINFO, _
                             ByVal wUsage As Long) As Long

Private Declare Function GetNearestColor _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal crColor As Long) As Long

Private Declare Function SetDIBitsToDevice _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long, _
                             ByVal dx As Long, _
                             ByVal dy As Long, _
                             ByVal SrcX As Long, _
                             ByVal SrcY As Long, _
                             ByVal Scan As Long, _
                             ByVal NumScans As Long, _
                             Bits As Any, _
                             BitsInfo As BITMAPINFO, _
                             ByVal wUsage As Long) As Long

Private Declare Function MoveToEx _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long, _
                             lpPoint As POINT) As Long

Private Declare Function LineTo _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function CreatePen _
                Lib "gdi32" (ByVal nPenStyle As Long, _
                             ByVal nWidth As Long, _
                             ByVal crColor As Long) As Long

Private Declare Function FillRect _
                Lib "user32" (ByVal hDC As Long, _
                              lpRect As RECT, _
                              ByVal hBrush As Long) As Long

Private Enum DrawTextFlags

    DT_TOP = &H0
    DT_LEFT = &H0
    DT_CENTER = &H1
    DT_RIGHT = &H2
    DT_VCENTER = &H4
    DT_BOTTOM = &H8
    DT_WORDBREAK = &H10
    DT_SINGLELINE = &H20
    DT_EXPANDTABS = &H40
    DT_TABSTOP = &H80
    DT_NOCLIP = &H100
    DT_EXTERNALLEADING = &H200
    DT_CALCRECT = &H400
    DT_NOPREFIX = &H800
    DT_INTERNAL = &H1000
    DT_EDITCONTROL = &H2000
    DT_PATH_ELLIPSIS = &H4000
    DT_END_ELLIPSIS = &H8000
    DT_MODIFYSTRING = &H10000
    DT_RTLREADING = &H20000
    DT_WORD_ELLIPSIS = &H40000
    DT_NOFULLWIDTHCHARBREAK = &H80000
    DT_HIDEPREFIX = &H100000
    DT_PREFIXONLY = &H200000

End Enum

Private Enum isState

    statenormal = &H1
    stateHot = &H2
    statePressed = &H3
    statedisabled = &H4
    stateDefaulted = &H5

End Enum

Private Type RGBTRIPLE

    rgbBlue  As Byte
    rgbGreen As Byte
    rgbRed   As Byte

End Type

Private Const DST_TEXT        As Long = &H1

Private Const DST_PREFIXTEXT  As Long = &H2

Private Const DST_COMPLEX     As Long = &H0

Private Const DST_ICON        As Long = &H3

Private Const DST_BITMAP      As Long = &H4

Private Const DSS_NORMAL      As Long = &H0

Private Const DSS_UNION       As Long = &H10

Private Const DSS_DISABLED    As Long = &H20

Private Const DSS_MONO        As Long = &H80

Private Const DSS_RIGHT       As Long = &H8000

Private Const BDR_RAISEDOUTER As Long = &H1

Private Const BDR_SUNKENOUTER As Long = &H2

Private Const BDR_RAISEDINNER As Long = &H4

Private Const BDR_SUNKENINNER As Long = &H8

Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)

Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

Private Const BF_LEFT   As Long = &H1

Private Const BF_TOP    As Long = &H2

Private Const BF_RIGHT  As Long = &H4

Private Const BF_BOTTOM As Long = &H8

Private Const BF_RECT   As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const DefGradientColor1 = "&H00E0E0E0"

Private Const DefGradientColor2 = "&H00404040"

Private Const DefAngle = "125"

Private Const DefEnabled = True

Private Const DefStretch = False

Private m_BackColor       As OLE_COLOR

Private m_StdPicture      As StdPicture

Private m_Style           As eStyle

Private iStyleIconOffset  As Long

Private m_btnRect         As RECT

Private m_txtRect         As RECT

Private m_lRegion         As Long

Private m_CaptionAlign    As isbAlign

Private m_sCaption        As String

Private m_Font            As StdFont

Private m_lFontColor      As Long

Private m_iState          As isState

Private m_shadowText      As Boolean

Private m_GradientColor1  As OLE_COLOR

Private m_GradientColor2  As OLE_COLOR

Private m_Angle           As Integer

Private m_Enabled         As Boolean

Private m_Stretch         As Boolean

Private m_GradientButton  As Boolean

Private m_StdPictureAlign As isbAlign

Private m_icon            As StdPicture

Private m_UseMaskColor    As Boolean

Private m_StdPictureSize  As Long

Private m_MaskColor       As Long

Private lwFontAlign       As Long

Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Event Click()

Public Property Get Angle() As Integer

    Angle = m_Angle

End Property

Public Property Let Angle(ByVal vValue As Integer)

    m_Angle = vValue

    Call refresh
    PropertyChanged "Angle"

End Property

Public Property Get BackColor() As OLE_COLOR
   
    BackColor = m_BackColor
   
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
   
    m_BackColor = vNewValue
    PropertyChanged "BackColor"
    Call refresh
   
End Property

Private Sub BuildRegion()

    On Error GoTo BuildRegion_Error

    If m_lRegion Then DeleteObject m_lRegion

    m_lRegion = CreateRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight)

    SetWindowRgn UserControl.hwnd, m_lRegion, True
    Exit Sub

BuildRegion_Error:

End Sub

Public Property Get Caption() As String

    On Error GoTo Caption_Error

    Caption = m_sCaption
    Exit Property

Caption_Error:

End Property

Public Property Let Caption(ByVal NewCaption As String)
    ' Description: this is the "Caption" property.

    On Error GoTo Caption_Error

    m_sCaption = NewCaption
    PropertyChanged "Caption"
    UserControl_Resize
    refresh
    Exit Property

Caption_Error:

End Property

Public Property Get CaptionAlign() As isbAlign

    On Error GoTo CaptionAlign_Error

    CaptionAlign = m_CaptionAlign
    Exit Property

CaptionAlign_Error:

End Property

Public Property Let CaptionAlign(ByVal NewCaptionAlign As isbAlign)
    ' Description: this is the "CaptionAlign" property.

    On Error GoTo CaptionAlign_Error

    m_CaptionAlign = NewCaptionAlign
    PropertyChanged "CaptionAlign"
    UserControl_Resize
    'UserControl_Paint
    Exit Property

CaptionAlign_Error:

End Property

Private Sub DrawCaption()

    On Error GoTo DrawCaption_Error

    Dim lcolor    As Long

    Dim ltmpColor As Long

    Dim R         As RECT

    If m_shadowText And m_Enabled Then
        R = m_txtRect

        R.Left = m_txtRect.Left + 2
        R.Top = m_txtRect.Top + 2
        R.bottom = m_txtRect.bottom + 2
        R.Right = m_txtRect.Right + 2

        UserControl.ForeColor = vbBlack
        DrawText UserControl.hDC, m_sCaption, -1, R, lwFontAlign

    End If

    lcolor = m_lFontColor

    If Not m_Enabled Then
        lcolor = GetSysColor(COLOR_BTNFACE)

    End If

    ltmpColor = UserControl.ForeColor
    UserControl.ForeColor = lcolor
    DrawText UserControl.hDC, m_sCaption, -1, m_txtRect, lwFontAlign
    UserControl.ForeColor = ltmpColor
    Exit Sub

DrawCaption_Error:

End Sub

Private Sub DrawCtlEdge(hDC As Long, _
                        ByVal X As Single, _
                        ByVal Y As Single, _
                        ByVal w As Single, _
                        ByVal h As Single, _
                        Optional Style As Long = EDGE_RAISED, _
                        Optional ByVal Flags As Long = BF_RECT)

    On Error GoTo DrawCtlEdge_Error

    Dim R As RECT

    With R
        .Left = X
        .Top = Y
        .Right = X + w
        .bottom = Y + h

    End With

    DrawEdge hDC, R, Style, Flags
    Exit Sub

DrawCtlEdge_Error:

End Sub

Public Property Get Enabled() As Boolean

    Enabled = m_Enabled

End Property

Public Property Let Enabled(ByVal vData As Boolean)

    m_Enabled = vData
    'm_UseMaskColor = Not vData
    UserControl.Enabled = m_Enabled
    Call refresh
    PropertyChanged "Enabled"

End Property

Public Property Get Stretch() As Boolean

    Stretch = m_Stretch

End Property

Public Property Let Stretch(ByVal vData As Boolean)

    m_Stretch = vData
    Call refresh
    PropertyChanged "Stretch"

End Property

Public Property Let Style(eVal As eStyle)

    If eVal <> m_Style Then
        m_Style = eVal
        PropertyChanged "Style"
        
        refresh

    End If

End Property

Public Property Get Style() As eStyle
    Style = m_Style

End Property

Public Property Get Font() As StdFont

    On Error GoTo Font_Error

    Set Font = UserControl.Font
    Exit Property

Font_Error:

End Property

Public Property Set Font(newFont As StdFont)

    On Error GoTo Font_Error

    Set m_Font = newFont
    Set UserControl.Font = newFont
    refresh
    PropertyChanged "Font"
    Exit Property

Font_Error:

End Property

Public Property Get FontColor() As OLE_COLOR

    On Error GoTo FontColor_Error

    FontColor = m_lFontColor
    Exit Property

FontColor_Error:

End Property

Public Property Let FontColor(lFontColor As OLE_COLOR)

    On Error GoTo FontColor_Error

    m_lFontColor = lFontColor
    PropertyChanged "FontColor"
    refresh
    Exit Property

FontColor_Error:

End Property

Public Property Get GradientButton() As Boolean

    On Error GoTo GradientButton_Error

    GradientButton = m_GradientButton
    Exit Property

GradientButton_Error:

End Property

Public Property Let GradientButton(newValue As Boolean)

    On Error GoTo GradientButton_Error

    m_GradientButton = newValue
    PropertyChanged "GradientButton"
    refresh
    Exit Property

GradientButton_Error:

End Property

Public Property Get GradientColor1() As OLE_COLOR

    GradientColor1 = m_GradientColor1

End Property

Public Property Let GradientColor1(ByVal vData As OLE_COLOR)

    m_GradientColor1 = vData

    Call refresh
    PropertyChanged "GradientColor1"

End Property

Public Property Get GradientColor2() As OLE_COLOR

    GradientColor2 = m_GradientColor2

End Property

Public Property Let GradientColor2(ByVal vData As OLE_COLOR)

    m_GradientColor2 = vData

    Call refresh
    PropertyChanged "GradientColor2"

End Property

Public Property Get Picture() As StdPicture

    Set Picture = m_StdPicture

End Property

Public Property Set Picture(Value As StdPicture)
    Set m_StdPicture = Value
    PropertyChanged "Picture"

    'Call DrawIconWCaption
    refresh

End Property

Public Property Get PictureAlign() As isbAlign

    On Error GoTo PictureAlign_Error

    PictureAlign = m_StdPictureAlign
    Exit Property

PictureAlign_Error:

End Property

Public Property Let PictureAlign(ByVal NewIconAlign As isbAlign)

    On Error GoTo PictureAlign_Error

    m_StdPictureAlign = NewIconAlign
    PropertyChanged "PictureAlign"
    refresh
    Exit Property

PictureAlign_Error:

End Property

Private Sub refresh()

    'On Error Resume Next
    Dim I  As Long

    Dim r1 As Long

    Dim g1 As Long

    Dim b1 As Long

    Dim r2 As Long

    Dim g2 As Long

    Dim b2 As Long

    Dim uH As Long

    Dim uW As Long

    If m_Style = XP_Button Then
        DrawXPButton
    Else
  
        UserControl.Cls
  
        Call PaintGradient(UserControl.hDC, 0, 0, 10, 10, m_GradientColor1, m_GradientColor2, 0)
  
        UserControl.refresh
  
        c1 = GetPixel(UserControl.hDC, 0, 0)
        c2 = GetPixel(UserControl.hDC, 1, 0)
        'c3 = GetPixel(UserControl.hDC, 2, 0)
        c4 = GetPixel(UserControl.hDC, 3, 0)
  
        c10 = GetPixel(UserControl.hDC, 7, 0)
        c11 = GetPixel(UserControl.hDC, 8, 0)
        c12 = GetPixel(UserControl.hDC, 9, 0)
        c13 = GetPixel(UserControl.hDC, 10, 0)
  
        UserControl.Cls
  
        uH = ScaleHeight - 1
        uW = ScaleWidth - 1
  
        On Error GoTo 0
  
        If m_StdPicture Is Nothing Then
            Call PaintGradient(UserControl.hDC, 0, 0, uW, uH, m_GradientColor1, m_GradientColor2, m_Angle)
  
            Line (0, 0)-(uW, 0), RGB(41, 48, 54)
            Line (0, 0)-(0, uH), RGB(41, 48, 54)
            Line (uW, 0)-(uW, uH), RGB(52, 51, 49)
            Line (0, uH)-(uW, uH), RGB(14, 29, 32)
  
            Line (1, 1)-(uW - 1, 1), c2
            Line (1, 1)-(1, uH - 1), c2
            Line (uW - 1, 1)-(uW - 1, uH - 1), c12
            Line (1, uH - 1)-(uW - 1, uH - 1), c12
  
            Line (2, 2)-(uW - 2, 2), c3
            Line (2, 2)-(2, uH - 2), c3
            Line (uW - 2, 2)-(uW - 2, uH - 2), c11
            Line (2, uH - 2)-(uW - 2, uH - 2), c11
  
            Line (3, 3)-(uW - 3, 3), c4
            Line (3, 3)-(3, uH - 3), c4
            Line (uW - 3, 3)-(uW - 3, uH - 3), c10
            Line (3, uH - 3)-(uW - 3, uH - 3), c10
  
            Line (4, 4)-(4, uH - 4), c3
            Line (4, 4)-(uW - 4, 4), c3
  
        ElseIf GradientButton = True Then
            Call PaintGradient(UserControl.hDC, 0, 0, uW, uH, m_GradientColor1, m_GradientColor2, m_Angle)
  
            Line (0, 0)-(uW, 0), RGB(41, 48, 54)
            Line (0, 0)-(0, uH), RGB(41, 48, 54)
            Line (uW, 0)-(uW, uH), RGB(52, 51, 49)
            Line (0, uH)-(uW, uH), RGB(14, 29, 32)
  
            Line (1, 1)-(uW - 1, 1), c2
            Line (1, 1)-(1, uH - 1), c2
            Line (uW - 1, 1)-(uW - 1, uH - 1), c12
            Line (1, uH - 1)-(uW - 1, uH - 1), c12
  
            Line (2, 2)-(uW - 2, 2), c3
            Line (2, 2)-(2, uH - 2), c3
            Line (uW - 2, 2)-(uW - 2, uH - 2), c11
            Line (2, uH - 2)-(uW - 2, uH - 2), c11
  
            Line (3, 3)-(uW - 3, 3), c4
            Line (3, 3)-(3, uH - 3), c4
            Line (uW - 3, 3)-(uW - 3, uH - 3), c10
            Line (3, uH - 3)-(uW - 3, uH - 3), c10
  
            Line (4, 4)-(4, uH - 4), c3
            Line (4, 4)-(uW - 4, 4), c3
        Else
     
            UserControl.BackColor = m_BackColor
      
            Line (0, 0)-(uW, 0), RGB(41, 48, 54)
            Line (0, 0)-(0, uH), RGB(41, 48, 54)
            Line (uW, 0)-(uW, uH), RGB(52, 51, 49)
            Line (0, uH)-(uW, uH), RGB(14, 29, 32)
  
            Line (1, 1)-(uW - 1, 1), &H8000000C
            Line (1, 1)-(1, uH - 1), &H8000000C
            Line (uW - 1, 1)-(uW - 1, uH - 1), &H8000000C
            Line (1, uH - 1)-(uW - 1, uH - 1), &H8000000C

        End If

    End If

    ''''''Draw picture
    If Not m_StdPicture Is Nothing Then
        If Picture <> 0 Then

            Dim ix As Long

            Dim iy As Long

            If Not m_StdPicture Is Nothing Then
                m_StdPictureSizeX = ScaleX(m_StdPicture.Width, vbHimetric, vbPixels)
                m_StdPictureSizeY = ScaleY(m_StdPicture.Height, vbHimetric, vbPixels)

            End If
      
            If m_StdPicture.Type = vbPicTypeIcon Or m_StdPicture.Type = vbPicTypeBitmap Then
                lFlags = DST_ICON
                Call DrawIconWCaption
                GoTo skip

            End If

            If m_StdPictureAlign = isbCenter Then
                ix = (UserControl.ScaleWidth - m_StdPictureSizeX) / 2
                iy = (UserControl.ScaleHeight - m_StdPictureSizeY) / 2
            ElseIf m_StdPictureAlign = isbbottom Then
                ix = (UserControl.ScaleWidth - m_StdPictureSizeX) / 2
                iy = UserControl.ScaleHeight - m_StdPictureSizeY - iStyleIconOffset
            ElseIf m_StdPictureAlign = isbTop Then
                ix = (UserControl.ScaleWidth - m_StdPictureSizeX) / 2
                iy = iStyleIconOffset
            ElseIf m_StdPictureAlign = isbleft Then
                ix = iStyleIconOffset
                iy = (UserControl.ScaleHeight - m_StdPictureSizeX) / 2
            ElseIf m_StdPictureAlign = isbRight Then
                ix = UserControl.ScaleWidth - m_StdPictureSizeX - iStyleIconOffset
                iy = (UserControl.ScaleHeight - m_StdPictureSizeY) / 2

            End If

            Dim ni As Long

            Dim nj As Long

            If m_iState = statePressed Then
                ix = ix + 1
                iy = iy + 1
            ElseIf m_iState = stateHot Then

                If m_iStyle = isbOfficeXP Then
                    If m_UseMaskColor Then
                        'This was added By t_eee eeee
                        TransBlt UserControl.hDC, ix + 1, iy + 1, m_StdPictureSizeX, m_StdPictureSizeY, m_StdPicture, m_MaskColor, &H808080
                        TransBlt UserControl.hDC, ix - 1, iy - 1, m_StdPictureSizeX, m_StdPictureSizeY, m_StdPicture, m_MaskColor '                        pMask.PaintPicture m_StdPicture,
                        '   ix, iy, m_StdPictureSize, m_StdPictureSize, , , , , vbSrcCopy
                    Else

                        PaintPicture m_StdPicture, ix, iy, m_StdPictureSizeX, m_StdPictureSizeY
                        lTransColor = GetPixel(UserControl.hDC, 1, 1)

                        For nj = iy To iy + m_StdPictureSize
                            For ni = ix To ix + m_StdPictureSize
                                lcurrpix = GetPixel(UserControl.hDC, ni, nj)

                                If lcurrpix <> lTransColor Then
                                    If m_UseMaskColor Then
                                        If lcurrpix <> m_MaskColor Then
                                            SetPixelV UserControl.hDC, ni, nj, &H808080

                                        End If

                                    Else
                                        SetPixelV UserControl.hDC, ni, nj, &H808080

                                    End If

                                End If

                            Next ni

                        Next nj

                    End If

                    ix = ix - 2
                    iy = iy - 2

                End If

            End If

            'I'll try to mask when usemaskcolor is true

            If m_UseMaskColor Then
                If m_Enabled Then

                    'Paint in the about pic on color
                    On Error GoTo MalformedIcon

                    TransBlt UserControl.hDC, ix, iy, m_StdPictureSizeX, m_StdPictureSizeY, m_StdPicture, m_MaskColor
                Else

                    'Disabled
                    On Error GoTo MalformedIcon

                    TransBlt UserControl.hDC, ix, iy, m_StdPictureSizeX, m_StdPictureSizeY, m_StdPicture, m_MaskColor, , , True

                End If

            Else
MalformedIcon:

                If m_Enabled Then
                    PaintPicture m_StdPicture, ix, iy, m_StdPictureSizeX, m_StdPictureSizeY
                Else
                    TransBlt UserControl.hDC, ix, iy, m_StdPictureSizeX, m_StdPictureSizeY, m_StdPicture, m_MaskColor, , , True

                End If

            End If

        End If

    End If
  
skip:
    DrawCaption

    Select Case m_iState

        Case statenormal, stateHot, stateDefaulted

            'DrawCtlEdge UserControl.hdc, 1, 1, UserControl.ScaleWidth, UserControl.ScaleHeight,
            '   BDR_RAISEDINNER
        Case statePressed
            'DrawCtlEdge UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight,
            '   BDR_SUNKENOUTER
            BitBlt UserControl.hDC, 1, 1, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3, UserControl.hDC, 0, 0, vbSrcCopy

    End Select

    UserControl.refresh

End Sub

Public Property Get ShadowText() As Boolean

    On Error GoTo Caption_Error

    ShadowText = m_shadowText
    Exit Property

Caption_Error:

End Property

Public Property Let ShadowText(ByVal vValue As Boolean)

    On Error GoTo Caption_Error

    m_shadowText = vValue
    PropertyChanged "ShadowText"
    UserControl_Resize
    refresh
    Exit Property

Caption_Error:

End Property

Private Sub TransBlt(ByVal DstDC As Long, _
                     ByVal DstX As Long, _
                     ByVal DstY As Long, _
                     ByVal DstW As Long, _
                     ByVal DstH As Long, _
                     ByVal SrcPic As StdPicture, _
                     Optional ByVal TransColor As Long = -1, _
                     Optional ByVal BrushColor As Long = -1, _
                     Optional ByVal MonoMask As Boolean = False, _
                     Optional ByVal isGreyscale As Boolean = False, _
                     Optional ByVal XPBlend As Boolean = False)

    If DstW = 0 Or DstH = 0 Then Exit Sub

    Dim b        As Long

    Dim h        As Long

    Dim f        As Long

    Dim I        As Long

    Dim newW     As Long

    Dim TmpDC    As Long

    Dim TmpBmp   As Long

    Dim TmpObj   As Long

    Dim Sr2DC    As Long

    Dim Sr2Bmp   As Long

    Dim Sr2Obj   As Long

    Dim Data1()  As RGBTRIPLE

    Dim Data2()  As RGBTRIPLE

    Dim Info     As BITMAPINFO

    Dim BrushRGB As RGBTRIPLE

    Dim gCol     As Long

    Dim SrcDC    As Long

    Dim tObj     As Long

    Dim ttt      As Long

    SrcDC = CreateCompatibleDC(hDC)

    If DstW < 0 Then DstW = UserControl.ScaleX(SrcPic.Width, 8, UserControl.ScaleMode)
    If DstH < 0 Then DstH = UserControl.ScaleY(SrcPic.Height, 8, UserControl.ScaleMode)

    If SrcPic.Type = 1 Then
        tObj = SelectObject(SrcDC, SrcPic)
    Else

        Dim hBrush As Long

        tObj = SelectObject(SrcDC, CreateCompatibleBitmap(DstDC, DstW, DstH))
        hBrush = CreateSolidBrush(MaskColor)
        DrawIconEx SrcDC, 0, 0, SrcPic.handle, 0, 0, 0, hBrush, &H1 Or &H2
        DeleteObject hBrush

    End If

    TmpDC = CreateCompatibleDC(SrcDC)
    Sr2DC = CreateCompatibleDC(SrcDC)
    TmpBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    Sr2Bmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    TmpObj = SelectObject(TmpDC, TmpBmp)
    Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)

    ReDim Data1(DstW * DstH * 3 - 1)
    ReDim Data2(UBound(Data1))

    With Info.bmiHeader
        .biSize = Len(Info.bmiHeader)
        .biWidth = DstW
        .biHeight = DstH
        .biPlanes = 1
        .biBitCount = 24

    End With

    BitBlt TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
    BitBlt Sr2DC, 0, 0, DstW, DstH, SrcDC, 0, 0, vbSrcCopy
    GetDIBits TmpDC, TmpBmp, 0, DstH, Data1(0), Info, 0
    GetDIBits Sr2DC, Sr2Bmp, 0, DstH, Data2(0), Info, 0

    If BrushColor > 0 Then
        BrushRGB.rgbBlue = (BrushColor \ &H10000) Mod &H100
        BrushRGB.rgbGreen = (BrushColor \ &H100) Mod &H100
        BrushRGB.rgbRed = BrushColor And &HFF

    End If

    If Not m_UseMaskColor Then TransColor = -1

    newW = DstW - 1

    For h = 0 To DstH - 1
        f = h * DstW

        For b = 0 To newW
            I = f + b

            If GetNearestColor(hDC, CLng(Data2(I).rgbRed) + 256& * Data2(I).rgbGreen + 65536 * Data2(I).rgbBlue) <> TransColor Then

                With Data1(I)

                    If BrushColor > -1 Then
                        If MonoMask Then
                            If (CLng(Data2(I).rgbRed) + Data2(I).rgbGreen + Data2(I).rgbBlue) <= 384 Then Data1(I) = BrushRGB
                        Else
                            Data1(I) = BrushRGB

                        End If

                    Else

                        If isGreyscale Then
                            gCol = CLng(Data2(I).rgbRed * 0.3) + Data2(I).rgbGreen * 0.59 + Data2(I).rgbBlue * 0.11
                            .rgbRed = gCol: .rgbGreen = gCol: .rgbBlue = gCol
                        Else

                            If XPBlend Then
                                .rgbRed = (CLng(.rgbRed) + Data2(I).rgbRed * 2) \ 3
                                .rgbGreen = (CLng(.rgbGreen) + Data2(I).rgbGreen * 2) \ 3
                                .rgbBlue = (CLng(.rgbBlue) + Data2(I).rgbBlue * 2) \ 3
                            Else
                                Data1(I) = Data2(I)

                            End If

                        End If

                    End If

                End With

            End If

        Next b

    Next h

    SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data1(0), Info, 0

    Erase Data1, Data2
    DeleteObject SelectObject(TmpDC, TmpObj)
    DeleteObject SelectObject(Sr2DC, Sr2Obj)

    If SrcPic.Type = 3 Then DeleteObject SelectObject(SrcDC, tObj)
    DeleteDC TmpDC: DeleteDC Sr2DC
    DeleteObject tObj: DeleteDC SrcDC

End Sub

Private Function TranslateColor(ByVal lcolor As Long) As Long
    'System color code to long rgb

    On Error GoTo TranslateColor_Error

    If OleTranslateColor(lcolor, 0, TranslateColor) Then
        TranslateColor = -1

    End If

    Exit Function

TranslateColor_Error:

End Function

Public Property Get UseMaskColor() As Boolean

    On Error GoTo UseMaskColor_Error

    UseMaskColor = m_UseMaskColor
    Exit Property

UseMaskColor_Error:

End Property

Public Property Let UseMaskColor(newValue As Boolean)

    On Error GoTo UseMaskColor_Error

    m_UseMaskColor = newValue
    PropertyChanged "UseMaskColor"
    refresh
    Exit Property

UseMaskColor_Error:

End Property

Private Sub UserControl_InitProperties()
    lwFontAlign = DT_CENTER Or DT_WORDBREAK 'DT_VCENTER Or DT_CENTER
    m_Enabled = True
  
    m_CaptionAlign = isbCenter
    m_StdPictureAlign = isbleft
    iStyleIconOffset = 20

End Sub

Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

    If Button = vbLeftButton Then
        m_iState = statePressed
        refresh

    End If

    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

    On Error GoTo UserControl_MouseUp_Error

    If Button = vbLeftButton Then

        m_iState = statenormal
        DoEvents

    End If

    refresh
    RaiseEvent MouseUp(Button, Shift, X, Y)
  
    If Button = vbLeftButton Then RaiseEvent Click
  
    Exit Sub

UserControl_MouseUp_Error:

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_iStyle = PropBag.ReadProperty("Style", 3)
    m_StdPictureAlign = PropBag.ReadProperty("PictureAlign", isbleft)
    m_sCaption = PropBag.ReadProperty("Caption", "")
    m_CaptionAlign = PropBag.ReadProperty("CaptionAlign", isbCenter)
    m_lFontColor = PropBag.ReadProperty("FontColor", GetSysColor(COLOR_BTNTEXT))
    Set UserControl.Font = PropBag.ReadProperty("Font", UserControl.Font)
    m_shadowText = PropBag.ReadProperty("ShadowText", False)
    m_Angle = PropBag.ReadProperty("Angle", DefAngle)
    m_GradientColor1 = PropBag.ReadProperty("GradientColor1", DefGradientColor1)
    m_GradientColor2 = PropBag.ReadProperty("GradientColor2", DefGradientColor1)
    Set m_StdPicture = PropBag.ReadProperty("Picture", Nothing)
    Enabled = PropBag.ReadProperty("Enabled", DefEnabled)
    Stretch = PropBag.ReadProperty("Stretch", DefStretch)
    GradientButton = PropBag.ReadProperty("GradientButton", True)
    m_UseMaskColor = PropBag.ReadProperty("UseMaskColor", False)
    m_MaskColor = PropBag.ReadProperty("MaskColor", &HC0C0C0)
    m_BackColor = PropBag.ReadProperty("BackColor", UserControl.Parent.BackColor)
    m_Style = PropBag.ReadProperty("Style", 1)

End Sub

Private Sub UserControl_Resize()

    Dim tmpRect As RECT

    Dim lH      As Long

    Dim lW      As Long

    On Error Resume Next

    If UserControl.Width < 300 Then UserControl.Width = 300
    If UserControl.Height < 300 Then UserControl.Height = 300

    lH = UserControl.ScaleHeight
    lW = UserControl.ScaleWidth
    SetRect m_btnRect, 0, 0, lW, lH
    SetRect m_txtRect, 0, 0, lW, lH
    CopyRect tmpRect, m_txtRect

    DrawText UserControl.hDC, m_sCaption, Len(m_sCaption), tmpRect, DT_CALCRECT Or DT_WORDBREAK Or IIf(m_bRTLText, DT_RTLREADING, 0)

    Select Case m_CaptionAlign

        Case isbCenter
            SetRect m_txtRect, (lW - tmpRect.Right - tmpRect.Left) / 2, (lH - tmpRect.bottom - tmpRect.Top) / 2, (lW + tmpRect.Right - tmpRect.Left) / 2 + 3, (lH + tmpRect.bottom - tmpRect.Top) / 2
            lwFontAlign = DT_CENTER Or DT_VCENTER Or DT_WORDBREAK

        Case isbleft
            CopyRect m_txtRect, tmpRect
            SetRect m_txtRect, iStyleIconOffset, (lH - tmpRect.bottom - tmpRect.Top) / 2, tmpRect.Right + iStyleIconOffset, (lH + tmpRect.bottom - tmpRect.Top) / 2 + 3
            lwFontAlign = DT_VCENTER Or DT_LEFT Or DT_WORDBREAK

        Case isbRight
            CopyRect m_txtRect, tmpRect
            SetRect m_txtRect, (lW - tmpRect.Right - tmpRect.Left) - iStyleIconOffset, (lH - tmpRect.bottom - tmpRect.Top) / 2, (lW - tmpRect.Left) - iStyleIconOffset, (lH + tmpRect.bottom - tmpRect.Top) / 2
            lwFontAlign = DT_VCENTER Or DT_RIGHT Or DT_WORDBREAK

        Case isbTop
            CopyRect m_txtRect, tmpRect
            SetRect m_txtRect, (lW - tmpRect.Right - tmpRect.Left) / 2, iStyleIconOffset / 2, (lW + tmpRect.Right - tmpRect.Left) / 2, iStyleIconOffset / 2 + (tmpRect.bottom - tmpRect.Top)
            lwFontAlign = DT_CENTER Or DT_TOP Or DT_WORDBREAK

        Case isbbottom
            CopyRect m_txtRect, tmpRect
            SetRect m_txtRect, (lW - tmpRect.Right - tmpRect.Left) / 2, lH - (tmpRect.bottom - tmpRect.Top) - iStyleIconOffset / 2, (lW + tmpRect.Right - tmpRect.Left) / 2 + 3, lH - iStyleIconOffset / 2
            lwFontAlign = DT_CENTER Or DT_BOTTOM Or DT_WORDBREAK

    End Select

    lwFontAlign = lwFontAlign Or IIf(m_bRTLText, DT_RTLREADING, 0)

    BuildRegion

    refresh

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag

        Call .WriteProperty("PictureAlign", m_StdPictureAlign, isbleft)
        Call .WriteProperty("Caption", m_sCaption, "")
        Call .WriteProperty("CaptionAlign", m_CaptionAlign, isbCenter)
        Call .WriteProperty("FontColor", m_lFontColor, GetSysColor(COLOR_BTNTEXT))
        Call .WriteProperty("Font", UserControl.Font)
        Call .WriteProperty("ShadowText", m_shadowText)
        Call .WriteProperty("Angle", m_Angle, DefAngle)
        Call .WriteProperty("GradientColor1", m_GradientColor1, DefGradientColor1)
        Call .WriteProperty("GradientColor2", m_GradientColor2, DefGradientColor2)
        Call .WriteProperty("BackColorTop", MyBackColorTop, DefBackColorTop)
        Call .WriteProperty("Picture", m_StdPicture, Nothing)
        Call .WriteProperty("Enabled", m_Enabled, DefEnabled)
        Call .WriteProperty("Stretch", m_Stretch, DefStretch)
        Call .WriteProperty("GradientButton", m_GradientButton, True)
        Call .WriteProperty("UseMaskColor", m_UseMaskColor, False)
        Call .WriteProperty("MaskColor", m_MaskColor, &HC0C0C0)
        Call .WriteProperty("BackColor", m_BackColor)
        Call .WriteProperty("Style", m_Style, 1)

    End With

End Sub

Private Function DrawXPButton()

    Dim I  As Long

    Dim r1 As Long, g1 As Long, b1 As Long

    Dim r2 As Long, g2 As Long, b2 As Long

    Dim uH As Long, uW As Long

    uH = ScaleHeight - 1
    uW = ScaleWidth - 1

    On Error Resume Next

    Line (0, 0)-(uW, uH), Parent.BackColor, BF

    On Error GoTo 0

    r1 = 236: g1 = 235: b1 = 230
    r2 = 214: g2 = 208: b2 = 197

    For I = 0 To uH - 3
        Line (1, I)-(uW, I), RGB(r1 * (I / (uH - 3)) + 255 - (255 * (I / (uH - 3))), g1 * (I / (uH - 3)) + 255 - (255 * (I / (uH - 3))), b1 * (I / (uH - 3)) + 255 - (255 * (I / (uH - 3))))
    Next
    
    For I = 0 To 3
        Line (0, uH - 4 + I)-(uW, uH - 4 + I), RGB(r2 * (I / 3) + r1 - (r1 * (I / 3)), g2 * (I / 3) + g1 - (g1 * (I / 3)), b2 * (I / 3) + b1 - (b1 * (I / 3)))
    Next
    
    PSet (0, 1), RGB(122, 149, 168): PSet (1, 0), RGB(122, 149, 168)
    Line (0, 2)-(2, 0), RGB(37, 87, 131) '7617536
    Line (2, 0)-(uW - 2, 0), 7617536
    PSet (uW - 1, 0), RGB(122, 149, 168): PSet (uW, 1), RGB(122, 149, 168)
    Line (uW - 2, 0)-(uW, 2), RGB(37, 87, 131)  '7617536
    Line (uW, 2)-(uW, uH - 2), 7617536
    PSet (uW, uH - 1), RGB(122, 149, 168): PSet (uW - 1, uH), RGB(122, 149, 168)
    Line (uW, uH - 2)-(uW - 2, uH), RGB(37, 87, 131) ' 7617536
    Line (uW - 2, uH)-(2, uH), 7617536
    PSet (1, uH), RGB(122, 149, 168): PSet (0, uH - 1), RGB(122, 149, 168)
    Line (2, uH)-(0, uH - 2), RGB(37, 87, 131)  '7617536
    Line (0, uH - 2)-(0, 2), 7617536

End Function

Public Sub DrawIconWCaption()

    Dim pW     As Long, ph As Long, lW As Long, lH As Long

    Dim StartX As Long, StartY As Long, lBrush As Long, lFlags As Long

    Dim lTemp  As Long, XCoord As Long, YCoord As Long
    
    If Not m_StdPicture Is Nothing Then
        pW = ScaleX(m_StdPicture.Width, vbHimetric, vbPixels)
        ph = ScaleY(m_StdPicture.Height, vbHimetric, vbPixels)

        If m_Stretch = True Then

        End If

    End If
    
    If LenB(m_Caption) Then
        lW = TextWidth(m_Caption)
        lH = TextHeight(m_Caption)

    End If

    Select Case m_StdPictureAlign

        Case Is = isbCenter
            StartX = ((ScaleWidth - pW) \ 2) + 1
            StartY = (ScaleHeight - ph) \ 2 + 1
            XCoord = Abs(ScaleWidth \ 2 - lW \ 2)
            YCoord = Abs(ScaleHeight \ 2 - lH \ 2)
            
        Case Is = isbTop
            StartX = ((ScaleWidth - pW) \ 2) + 1
            StartY = 6 '(ScaleHeight - (pH + lH)) \ 2 + 1
            XCoord = Abs(ScaleWidth \ 2 - lW \ 2)
            YCoord = Abs(ScaleHeight \ 2 + ph \ 2 - lH \ 2)

        Case Is = isbbottom
            StartX = (ScaleWidth - pW) \ 2
            StartY = (ScaleHeight - (ph - lH)) \ 2 + 1
            XCoord = Abs(ScaleWidth \ 2 - lW \ 2)
            YCoord = Abs(ScaleHeight \ 2 - (ph + lH) \ 2)

        Case Is = isbleft
            StartX = 4
            StartY = (ScaleHeight - ph) \ 2 + 1
            XCoord = Abs(ScaleWidth \ 2 - lW \ 2)
            YCoord = Abs(ScaleHeight \ 2 - lH \ 2)

        Case Is = isbRight
            StartX = ScaleWidth - 4 - pW
            StartY = (ScaleHeight - ph) \ 2 + 1
            XCoord = Abs(ScaleWidth \ 2 - lW \ 2)
            YCoord = Abs(ScaleHeight \ 2 - lH \ 2)

    End Select

    If m_Enabled Then lFlags = DST_PREFIXTEXT Or DSS_NORMAL Else lFlags = DST_PREFIXTEXT Or DSS_DISABLED

    If LenB(m_Caption) Then Call DrawStateText(hDC, 0&, 0&, m_Caption, Len(m_Caption), XCoord, YCoord, 0&, 0&, lFlags)
    
    If Not m_StdPicture Is Nothing Then
        If m_StdPicture.Type = vbPicTypeBitmap Then
            lFlags = DST_BITMAP
        ElseIf m_StdPicture.Type = vbPicTypeIcon Then
            lFlags = DST_ICON

        End If

        If Not m_Enabled Then
            lFlags = lFlags Or DSS_DISABLED 'Draw disabled
        
        End If

        With m_StdPicture
            DrawState hDC, lBrush, 0, .handle, 0, CLng(StartX), CLng(StartY), .Width, .Height, lFlags

        End With

    End If
    
End Sub
