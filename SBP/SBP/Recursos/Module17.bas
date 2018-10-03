Attribute VB_Name = "Module17"
'================================================
' Module:        mGradient.bas (any angle)
' Author:        Carles P.V. - 2005
' Dependencies:  None
' Last revision: 2005.05.21
'================================================
'
' 2005.05.21: Minor speed improvements
'             Thanks to Robert Rayment.
'
'================================================
'
' For fastest renderings, see my other post:
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=60477
' Only for horizontal, vertical and both diagonals directions.

Option Explicit

Private Type BITMAPINFOHEADER

    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long

End Type

Private Const DIB_RGB_COLORS As Long = 0

Private Declare Function StretchDIBits _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long, _
                             ByVal dx As Long, _
                             ByVal dy As Long, _
                             ByVal SrcX As Long, _
                             ByVal SrcY As Long, _
                             ByVal wSrcWidth As Long, _
                             ByVal wSrcHeight As Long, _
                             lpBits As Any, _
                             lpBitsInfo As Any, _
                             ByVal wUsage As Long, _
                             ByVal dwRop As Long) As Long

'//

Private Const PI      As Single = 3.14159265358979

Private Const TO_DEG  As Single = 180 / PI

Private Const TO_RAD  As Single = PI / 180

Private Const INT_ROT As Long = 1000 ' Increase this value for more precision

Public Sub PaintGradient(ByVal hDC As Long, _
                         ByVal X As Long, _
                         ByVal Y As Long, _
                         ByVal Width As Long, _
                         ByVal Height As Long, _
                         ByVal Color1 As Long, _
                         ByVal Color2 As Long, _
                         ByVal Angle As Single)

    Dim uBIH      As BITMAPINFOHEADER

    Dim lBits()   As Long

    Dim lGrad()   As Long

    Dim lClr      As Long

    Dim r1        As Long

    Dim g1        As Long

    Dim b1        As Long

    Dim r2        As Long

    Dim g2        As Long

    Dim b2        As Long

    Dim dR        As Long

    Dim dG        As Long

    Dim dB        As Long

    Dim Scan      As Long

    Dim I         As Long

    Dim j         As Long

    Dim iIn       As Long

    Dim jIn       As Long

    Dim iEnd      As Long

    Dim jEnd      As Long

    Dim Offset    As Long

    Dim lQuad     As Long

    Dim AngleDiag As Single

    Dim AngleComp As Single

    Dim g         As Long

    Dim luSin     As Long

    Dim luCos     As Long

    '-- Minor check

    If (Width > 0 And Height > 0) Then

        '-- Right-hand [+] (ox=0º)
        Angle = -Angle + 90

        '-- Normalize to [0º;360º]
        Angle = Angle Mod 360

        If (Angle < 0) Then Angle = 360 + Angle

        '-- Get quadrant (0 - 3)
        lQuad = Angle \ 90

        '-- Normalize to [0º;90º]
        Angle = Angle Mod 90

        '-- Calc. gradient length ('distance')

        If (lQuad Mod 2 = 0) Then
            AngleDiag = Atn(Width / Height) * TO_DEG
        Else
            AngleDiag = Atn(Height / Width) * TO_DEG

        End If

        AngleComp = (90 - Abs(Angle - AngleDiag)) * TO_RAD
        Angle = Angle * TO_RAD
        g = Sqr(Width * Width + Height * Height) * Sin(AngleComp) 'Sinus theorem

        '-- Decompose colors

        If (lQuad > 1) Then
            lClr = Color1
            Color1 = Color2
            Color2 = lClr

        End If

        r1 = (Color1 And &HFF&)
        g1 = (Color1 And &HFF00&) \ 256
        b1 = (Color1 And &HFF0000) \ 65536
        r2 = (Color2 And &HFF&)
        g2 = (Color2 And &HFF00&) \ 256
        b2 = (Color2 And &HFF0000) \ 65536

        '-- Get color distances
        dR = r2 - r1
        dG = g2 - g1
        dB = b2 - b1

        '-- Size gradient-colors array
        ReDim lGrad(0 To g - 1)

        '-- Calculate gradient-colors
        iEnd = g - 1

        If (iEnd = 0) Then
            '-- Special case (1-pixel wide gradient)
            lGrad(0) = (b1 \ 2 + b2 \ 2) + 256 * (g1 \ 2 + g2 \ 2) + 65536 * (r1 \ 2 + r2 \ 2)
        Else

            For I = 0 To iEnd
                lGrad(I) = b1 + (dB * I) \ iEnd + 256 * (g1 + (dG * I) \ iEnd) + 65536 * (r1 + (dR * I) \ iEnd)
            Next I

        End If

        '-- Size DIB array
        ReDim lBits(Width * Height - 1) As Long

        '-- Render gradient DIB

        iEnd = Width - 1
        jEnd = Height - 1

        Select Case lQuad

            Case 0, 2

                luSin = Sin(Angle) * INT_ROT
                luCos = Cos(Angle) * INT_ROT
                Offset = 0
                Scan = Width

            Case 1, 3

                luSin = Sin(90 * TO_RAD - Angle) * INT_ROT
                luCos = Cos(90 * TO_RAD - Angle) * INT_ROT
                Offset = jEnd * Width
                Scan = -Width

        End Select

        jIn = 0
        iIn = 0

        For j = 0 To jEnd
            iIn = jIn

            For I = 0 To iEnd
                lBits(I + Offset) = lGrad(iIn \ INT_ROT)
                iIn = iIn + luSin
            Next I

            jIn = jIn + luCos
            Offset = Offset + Scan
        Next j

        '-- Define DIB header

        With uBIH
            .biSize = 40
            .biPlanes = 1
            .biBitCount = 32
            .biWidth = Width
            .biHeight = Height

        End With

        '-- Paint it!
        Call StretchDIBits(hDC, X, Y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy)

    End If

End Sub

