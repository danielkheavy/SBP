VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmlabel 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Etiquetas"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12825
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLabel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   12825
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox columna 
      Height          =   375
      Left            =   8160
      MaxLength       =   1
      TabIndex        =   23
      Text            =   "2"
      Top             =   6720
      Width           =   2175
   End
   Begin VB.TextBox cantidad 
      Height          =   375
      Left            =   8160
      MaxLength       =   4
      TabIndex        =   22
      Text            =   "1"
      Top             =   7080
      Width           =   2175
   End
   Begin VB.CommandButton cmdSettings 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Orientation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5280
      Width           =   1035
   End
   Begin VB.CommandButton cmdSettings 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label Font"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6240
      Width           =   1035
   End
   Begin VB.CommandButton cmdSettings 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5760
      Width           =   1035
   End
   Begin VB.CommandButton cmdInfo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Visualizar Ayuda"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6960
      Width           =   3915
   End
   Begin VB.TextBox txtData 
      Height          =   360
      Left            =   7200
      TabIndex        =   11
      ToolTipText     =   "Enter sample data in order of appearance in the Label Definition, separated by pipes, ex. ""1234|ABC|01/01/2002""."
      Top             =   720
      Width           =   4395
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   360
      Left            =   8160
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   7440
      Width           =   3315
   End
   Begin MSComDlg.CommonDialog cdlgDef 
      Left            =   4080
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "*.def|*.def"
      InitDir         =   "C:\"
   End
   Begin VB.CommandButton cmdLabel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Abrir"
      Height          =   435
      Index           =   4
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Print the Label."
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox txtLabelDef 
      Height          =   6555
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   240
      Width           =   6555
   End
   Begin VB.CommandButton cmdLabel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Imprimir"
      Height          =   435
      Index           =   3
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print the Label."
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdLabel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Grabar"
      Enabled         =   0   'False
      Height          =   435
      Index           =   2
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print the Label."
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdLabel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Salir"
      Height          =   435
      Index           =   0
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Transfer the Item."
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdLabel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Preview"
      Height          =   435
      Index           =   1
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print the Label."
      Top             =   720
      Width           =   975
   End
   Begin VB.PictureBox picPreview 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1507
      Left            =   7200
      ScaleHeight     =   2.54
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   7.62
      TabIndex        =   0
      Top             =   1440
      Width           =   4387
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Columnas"
      Height          =   315
      Left            =   7200
      TabIndex        =   24
      Top             =   6720
      Width           =   1035
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cantidad"
      Height          =   315
      Left            =   7200
      TabIndex        =   21
      Top             =   7080
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "s|o|l"
      Height          =   375
      Left            =   9720
      TabIndex        =   20
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label lblLabelFont 
      BackColor       =   &H00808080&
      Caption         =   "= s|f|Arial|8|true"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   19
      Top             =   6300
      Width           =   2535
   End
   Begin VB.Label lblLabelSize 
      BackColor       =   &H00808080&
      Caption         =   "= s|s|inch|1|3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   18
      Top             =   5820
      Width           =   2535
   End
   Begin VB.Label lblOrientation 
      BackColor       =   &H00808080&
      Caption         =   "= s|o|p"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   17
      Top             =   5340
      Width           =   1335
   End
   Begin VB.Label lblSampleData 
      BackColor       =   &H00808080&
      Caption         =   "Datos:"
      Height          =   315
      Left            =   7200
      TabIndex        =   12
      Top             =   420
      Width           =   4395
   End
   Begin VB.Label lblPrinter 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Impresora"
      Height          =   315
      Left            =   7200
      TabIndex        =   10
      Top             =   7440
      Width           =   1035
   End
   Begin VB.Label lblPreview 
      BackColor       =   &H00808080&
      Caption         =   "Preview:"
      Height          =   255
      Left            =   7200
      TabIndex        =   8
      Top             =   1200
      Width           =   4395
   End
   Begin VB.Label lblDef 
      BackColor       =   &H00808080&
      Caption         =   "Definiciones"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   3915
   End
   Begin VB.Line Line1 
      X1              =   6840
      X2              =   6840
      Y1              =   120
      Y2              =   7980
   End
End
Attribute VB_Name = "frmlabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private m_sDefFile  As String

Private m_bLoading  As Boolean

Private m_bPortrait As Boolean

Private Sub cmdInfo_Click()

    'frmInfo.Show 1
End Sub

Private Sub cmdLabel_Click(Index As Integer)

    Me.MousePointer = vbHourglass

    Dim found   As Integer

    'Dim oFSO As New FileSystemObject
    'Dim tsDef As TextStream
    Dim sTemp   As String

    Dim lResult As Long

    Dim I       As Long

    Dim sline() As String

    Dim bfound  As Boolean

    Dim n       As Integer

    Dim m       As Integer

    Dim posx    As Integer

    Dim posy    As Integer

    Dim sw      As Integer
    
    Select Case Index

        Case 0  'Exit
            foo3_Click
            'Unload Me
            
        Case 1  'Preview Label

            If Len(txtData) = 0 Then Exit Sub
            PreviewLabel txtLabelDef.Text, txtData.Text
            cmdLabel(3).Enabled = True
            
        Case 2  'Save
            Err.Clear

            On Error Resume Next

            'm_sDefFile
            With cdlgDef
                .DialogTitle = "Grabar"

                If LenB(m_sDefFile) > 0 Then
                    .FileName = m_sDefFile

                End If
                
                .ShowSave

            End With
            
            If Err.Number = 0 Then
                m_sDefFile = cdlgDef.FileName
                
                'Add .def extension if necessary
                If Right$(m_sDefFile, 4) <> ".def" Then
                    m_sDefFile = m_sDefFile & ".def"

                End If
                
                'Save the file
                lResult = vbYes

                'If existe_archivo(m_sDefFile) > 0 Then
                If existe_archivo(m_sDefFile) > 0 Then
                    'If oFSO.FileExists(m_sDefFile) Then
                    lResult = MsgBox("Ya existe archivo.  Desea Reescribir?", vbQuestion + vbYesNo, App.Title)

                End If
                
                If lResult = vbYes Then
                    cmdLabel(2).Enabled = False
                    sline = Split(txtLabelDef.Text, vbcrlf)
                    
                    If LenB(txtLabelDef.Text) > 0 Then
                        borrar_archivo m_sDefFile
                        found = grabar_etiqueta(m_sDefFile)
                        'Set tsDef = oFSO.CreateTextFile(m_sDefFile, True)
                        
                        'For i = 0 To UBound(sline)
                        '    If i < UBound(sline) Then
                        '        tsDef.WriteLine sline(i)
                        '    Else
                        '        tsDef.Write sline(i)
                        '    End If
                        'Next
                        
                        'tsDef.Close
                    End If

                End If

            End If
            
            On Error GoTo 0
            
        Case 3  'Print

            If Val(cantidad) < 0 Then Exit Sub
            If Len(txtData) = 0 Then Exit Sub
            
            If cboPrinter.Text = "File" Then
                picPreview.Picture = picPreview.Image
                SavePicture picPreview.Picture, globalpath & "\001d\06\zebra\sample.bmp"
                
                MsgBox "The label was saved as " & globalpath & "\001d\06\zebra\sample.bmp", vbInformation + vbOKOnly, "Label Maker"
                
            ElseIf cboPrinter.Text <> "None" Then
                bfound = False

                For I = 0 To Printers.count - 1

                    If Printers(I).DeviceName = cboPrinter.Text Then
                        Set Printer = Printers(I)
                        bfound = True
                        Exit For

                    End If

                Next
                
                If bfound Then
                    'Set orientation
                    Printer.Orientation = 1

                    If Not m_bPortrait Then
                        Printer.Orientation = 2

                    End If
                    
                    'Print label
                    sw = 0
                    n = 1
                    posx = 0
                    posy = 0
                    'picPreview.Picture = picPreview.Image
                    Printer.PrintQuality = vbPRPQHigh
                    Do
                        m = 1
                        posx = 100
                        posy = 100

                        If n > Val(cantidad) Then Exit Do

                        For m = 1 To Val(columna)
                            picPreview.Picture = picPreview.Image
                            'Printer.Print
                            'Printer.PaintPicture picPreview.Picture, posx, posy, picPreview.Width, picPreview.Height
                            'Printer.PaintPicture picPreview.Picture, posx, 0, picPreview.Width, picPreview.Height, 0, 0, picPreview.Width, picPreview.Height
                            'Printer.PaintPicture picPreview.Picture, posx, 0, Printer.Width, Printer.Height
                            'picPreview.AutoRedraw = True
                            Printer.PaintPicture picPreview.Picture, posx, posy
                            'picPreview.AutoRedraw = False
                            n = n + 1

                            If n > Val(cantidad) Then Exit For
                            posx = posx + 3100
                            sw = 1
                        Next m

                        'posx = 0
                        'posy = 0
                        'Printer.Print
                        Printer.EndDoc
                    Loop
                    'If sw = 0 Then
                    '   Printer.EndDoc
                    'End If
                Else
                    MsgBox "Unable to locate specified printer - " & cboPrinter.Text & ".  Please verify the printer settings.", vbExclamation + vbOKOnly, "Printing Label"

                End If
                
            Else
                MsgBox "Please select a printer.", vbExclamation + vbOKOnly, App.Title

            End If
            
        Case 4  'Open

            If cmdLabel(2).Enabled Then
                lResult = MsgBox("Desea Guardar Cambios en la definicion de la etiqueta actual?", vbQuestion + vbYesNoCancel, App.Title)
                
                If lResult = vbYes Then
                    cmdLabel_Click 2
                ElseIf lResult = vbCancel Then
                    GoTo cmdLabel_Click_EXIT
                Else

                    'Do nothing
                End If

            End If
            
            m_bLoading = True
            
            Err.Clear

            On Error Resume Next
            
            With cdlgDef
                .DialogTitle = "Abrir etiqueta"
                .InitDir = globalpath & "\001d\06\zebra"
                
                .ShowOpen

            End With
            
            If Err.Number = 0 Then
                sTemp = cdlgDef.FileName

                'MsgBox sTemp
                If existe_archivo(sTemp) > 0 Then
                    'If oFSO.FileExists(sTemp) Then
                    '    Set tsDef = oFSO.OpenTextFile(sTemp, ForReading, False, TristateFalse)
                    
                    'txtLabelDef.Text = tsDef.ReadAll
                    
                    'tsDef.Close
                    found = leer_etiqueta(sTemp)
                
                End If

            Else
                sTemp = ""

            End If
            
            On Error GoTo 0
            
            cmdLabel(2).Enabled = False
            m_sDefFile = sTemp
            m_bLoading = False
            
    End Select
    
cmdLabel_Click_EXIT:
    Erase sline
    'Set tsDef = Nothing
    'Set oFSO = Nothing
    
    Me.MousePointer = vbNormal
    
End Sub

Function grabar_etiqueta(buf As String)

    Dim I       As Integer

    Dim sline() As String

    sline = Split(txtLabelDef.Text, vbcrlf)
                
    Open buf For Append As #1

    For I = 0 To UBound(sline)

        If I < UBound(sline) Then
            'tsDef.WriteLine sLine(i)
            Print #1, sline(I)
        Else
            'tsDef.Write sLine(i)
            Print #1, sline(I)

        End If

    Next

    Close #1
 
End Function

Function leer_etiqueta(buf As String)

    On Error GoTo cmd19012_err

    Dim linea As String, cTotal As String

    txtLabelDef.Text = ""
    Open buf For Input As #1

    Do Until EOF(1)
        Line Input #1, linea
        cTotal = cTotal + linea + vbcrlf
    Loop
    Close #1
    txtLabelDef.Text = cTotal
    leer_etiqueta = 1
    Exit Function
cmd19012_err:
    Exit Function

End Function

Private Sub cmdSettings_Click(Index As Integer)
    Exit Sub

    Dim sTemp As String
    
    Select Case Index

        Case 0  'Insert Font
            Err.Clear

            On Error Resume Next
            
            With cdlgDef
                .ShowFont

            End With
            
            If Err.Number = 0 Then
                
            End If
            
            On Error GoTo 0
            
        Case 1  'Insert Label Size

            With frmLabelSize
                .Show vbModal
                
                If Not .Cancelled Then
                    
                End If

            End With
            
            Unload frmLabelSize
            
        Case 2  'Insert Orientation

            With frmOrientation
                .Show vbModal
                
                If Not .Cancelled Then
                    
                End If

            End With
            
            Unload frmOrientation
            
    End Select
    
    If Not m_bLoading Then
        cmdLabel(2).Enabled = True

    End If
    
End Sub

Private Sub dl8822_Click()

End Sub

Private Sub foo3_Click()

End Sub

Private Sub Form_Load()

    Dim I As Long

    On Error GoTo cmd9012_err

    'Load printer list
    cboPrinter.Clear
    cboPrinter.AddItem "None"
    cboPrinter.AddItem "File"
    Printer.Orientation = 1

    For I = 0 To Printers.count - 1
        cboPrinter.AddItem Printers(I).DeviceName
    Next
    cboPrinter.ListIndex = 0
    Exit Sub
cmd9012_err:
    MsgBox "Seleccione una impresora ", 48, "Aviso"
    Exit Sub
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim lResult As Long
    
    'If frmInfo.Visible Then
    '    Unload frmInfo
    'End If
    
    If cmdLabel(2).Enabled Then
        lResult = MsgBox("Desea Grabar los cambios ?", vbQuestion + vbYesNoCancel, "Salir etiqueta")

        If lResult = vbYes Then
            cmdLabel_Click 2
            'End
            foo3_Click
        ElseIf lResult = vbCancel Then
            Cancel = True
        Else
            foo3_Click

        End If

    Else
        'End
        foo3_Click

    End If
    
End Sub

Private Sub orientacion_Click()

End Sub

Private Sub fraSettings_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub picPreview_Resize()

    With picPreview

        If (.Left + .Width) > Me.Width Then
            Me.Width = .Left + .Width + 175

        End If

    End With
    
End Sub

Private Sub txtLabelDef_Change()

    If Not m_bLoading Then
        cmdLabel(2).Enabled = True

    End If
    
End Sub

Private Sub PreviewLabel(ByVal sLabelDef As String, ByVal sLineData As String)

    Me.MousePointer = vbHourglass
    
    Dim lResult     As Long

    Dim sline()     As String

    Dim sChunk()    As String

    Dim sSampData() As String

    Dim sBarcode    As String

    Dim I           As Long

    Dim j           As Long

    Dim iCalc       As Integer

    Dim sPrinter    As String

    Dim nVarCount   As Long

    Dim nStart      As Long

    Dim nPos        As Long

    Dim nCurVar     As Long

    Dim sVarName    As String

    Dim sWork()     As String

    Dim nSplit1     As Long

    Dim nSplit2     As Long

    Dim sTemp       As String

    Dim sText1      As String

    Dim sText2      As String

    Dim bI25        As Boolean 'Barcode Format = I25 or Code39
    
    'Initialize Format
    bI25 = False
    
    'Split Label definition and sample data
    sline = Split(sLabelDef, vbcrlf)
    sSampData = Split(sLineData, "|")
    
    'Initialize PictureBox - settings & dimensions
    picPreview.ScaleMode = vbPixels
    picPreview.Cls
    picPreview.Picture = Nothing
    picPreview.refresh
    Me.refresh
    
    'Process Label Definition
    nCurVar = 0

    For j = 0 To UBound(sline)

        If Not VerifyLineData(sline(j)) Then
            GoTo PreviewLabel_BAD_DATA

        End If
        
        sChunk = Split(sline(j), "|")
        
        nStart = 1

        Select Case sChunk(0)

            Case "s"    'Setting

                If sChunk(1) = "c" Then
                    If sChunk(2) = "I25" Then
                        'Use I25 Barcode Format
                        bI25 = True

                    End If

                ElseIf sChunk(1) = "o" Then

                    'Set portrait or landscape
                    If sChunk(2) = "p" Then
                        m_bPortrait = True
                    Else
                        m_bPortrait = False

                    End If

                ElseIf sChunk(1) = "s" Then

                    'Set Scale mode
                    If sChunk(2) = "in" Then
                        picPreview.ScaleMode = vbInches
                    ElseIf sChunk(2) = "pix" Then
                        picPreview.ScaleMode = vbPixels
                    Else 'If sChunk(2) = "cm" Then
                        picPreview.ScaleMode = vbCentimeters

                    End If
                    
                    'Set label size
                    With picPreview

                        'Set height
                        If .ScaleHeight > CInt(sChunk(3)) Then
                            Do
                                .Height = .Height - 1
                            Loop Until .ScaleHeight < CInt(sChunk(3))

                        End If
                        
                        Do
                            .Height = .Height + 1
                        Loop Until .ScaleHeight > CInt(sChunk(3))
                        
                        .Height = .Height - 1
                        
                        'Set Width
                        If .ScaleWidth > CInt(sChunk(4)) Then
                            Do
                                .Width = .Width - 1
                            Loop Until .ScaleWidth < CInt(sChunk(4))

                        End If
                        
                        Do
                            .Width = .Width + 1
                        Loop Until .ScaleWidth > CInt(sChunk(4))
                        
                        .Width = .Width - 1
                        
                        'Reset the Scale mode
                        .ScaleMode = vbPixels

                    End With
                    
                ElseIf sChunk(1) = "f" Then

                    'Set label font
                    With picPreview
                        .Font = sChunk(2)
                        .FontSize = CInt(sChunk(3))
                        .FontBold = CBool(sChunk(4))
                        .FontItalic = CBool(sChunk(5))

                    End With

                End If
                
            Case "f"    'Frame
                'No other information listed on this line
                'Top
                picPreview.Line (0, 0)-(picPreview.ScaleWidth - 1, 0), &H0&, BF
                'Bottom
                picPreview.Line (0, picPreview.ScaleHeight - 1)-(picPreview.ScaleWidth - 1, picPreview.ScaleHeight - 1), &H0&, BF
                'Left Side
                picPreview.Line (0, 0)-(0, picPreview.ScaleHeight - 1), &H0&, BF
                'Right Side
                picPreview.Line (picPreview.ScaleWidth - 1, 0)-(picPreview.ScaleWidth - 1, picPreview.ScaleHeight - 1), &H0&, BF
                
            Case "b"    'Barcode
                'Barcode only contains actual barcode number -- no special formatting required
                nPos = InStr(1, sChunk(3), "~")

                If nPos > 0 Then
                    nStart = nPos + 1
                    sVarName = Mid$(sChunk(3), nPos, InStr(nStart, sChunk(3), "~") - nPos + 1)
                    sChunk(3) = Replace(sChunk(3), sVarName, sSampData(nCurVar))
                    nCurVar = nCurVar + 1

                End If
                
                'Generate Barcode
                If bI25 Then

                    'MsgBox "xxx"
                    Dim oI25Barcode As New clsI25Barcode

                    lResult = oI25Barcode.GenerateI25Barcode(picPreview, sChunk(3), sChunk(1), CLng(sChunk(2)))
                    Set oI25Barcode = Nothing
                Else

                    Dim oBarcode As New clsBarcode

                    lResult = oBarcode.GenerateBarCode(picPreview, sChunk(3), sChunk(1), CInt(sChunk(2)))
                    Set oBarcode = Nothing

                End If
                
            Case "t"    'Text
                'Print text on label, where a value of "c" for the x position
                'means center the entry on the label
                nPos = InStr(1, sChunk(3), "~")

                Do Until nPos = 0
                    nStart = nPos + 1
                    sVarName = Mid$(sChunk(3), nPos, InStr(nStart, sChunk(3), "~") - nPos + 1)
                    sChunk(3) = Replace(sChunk(3), sVarName, sSampData(nCurVar))
                    nCurVar = nCurVar + 1
                    
                    nStart = InStr(nStart, sChunk(3), "~") + 1
                    nPos = InStr(nStart, sChunk(3), "~")
                Loop
                
                'Print Text
                With picPreview

                    If sChunk(1) = "c" Then
                        iCalc = (.ScaleWidth - .TextWidth(sChunk(3))) / 2
                        .CurrentX = IIf(iCalc > 0, iCalc, 0)
                    ElseIf CLng(sChunk(1)) < 0 Then
                        .CurrentX = .ScaleWidth - .TextWidth(sChunk(3)) + CLng(sChunk(1))
                    Else 'If CLng(sChunk(1)) >= 0 Then
                        .CurrentX = CLng(sChunk(1))

                    End If

                    .CurrentY = CLng(sChunk(2))

                End With
                
                picPreview.Print sChunk(3)
                
            Case "d"    'Double-line Text
                'Print up to 2 lines of text on label, where a value of "c"
                'for the x position means center the entry on the label.
                nPos = InStr(1, sChunk(3), "~")

                Do Until nPos = 0
                    nStart = nPos + 1
                    sVarName = Mid$(sChunk(3), nPos, InStr(nStart, sChunk(3), "~") - nPos + 1)
                    sChunk(3) = Replace(sChunk(3), sVarName, sSampData(nCurVar))
                    nCurVar = nCurVar + 1
                    
                    nStart = InStr(nStart, sChunk(3), "~") + 1
                    nPos = InStr(nStart, sChunk(3), "~")
                Loop
                
                'Chunk up text
                sWork = Split(sChunk(3), " ")
                
                '****************
                '* First Line
                '****************
                'Initialize
                nSplit1 = 0
                sTemp = ""
                sText1 = ""
                
                'Determine split points
                For I = 0 To UBound(sWork)
                    sTemp = sTemp & IIf(I > 0, " ", "") & sWork(I)

                    If picPreview.TextWidth(sTemp) > picPreview.ScaleWidth Then
                        nSplit1 = I - 1
                        Exit For

                    End If

                Next

                If nSplit1 = 0 Then
                    nSplit1 = UBound(sWork)

                End If

                For I = 0 To nSplit1
                    sText1 = sText1 & IIf(I > 0, " ", "") & sWork(I)
                Next
                
                'Print Text
                With picPreview

                    If sChunk(1) = "c" Then
                        iCalc = (.ScaleWidth - .TextWidth(sText1)) / 2
                        .CurrentX = IIf(iCalc > 0, iCalc, 0)
                    ElseIf CLng(sChunk(1)) < 0 Then
                        .CurrentX = .ScaleWidth - .TextWidth(sText1) + CLng(sChunk(1))
                    Else 'If CLng(sChunk(1)) >= 0 Then
                        .CurrentX = CLng(sChunk(1))

                    End If

                    .CurrentY = CLng(sChunk(2))

                End With
                
                picPreview.Print sText1
                
                '****************
                '* Second Line
                '****************
                If nSplit1 < UBound(sWork) Then
                    'Initialize
                    nSplit1 = nSplit1 + 1
                    nSplit2 = 0
                    sTemp = ""
                    sText2 = ""
                    
                    'Determine split points
                    For I = nSplit1 To UBound(sWork)
                        sTemp = sTemp & IIf(I > nSplit1, " ", "") & sWork(I)

                        If picPreview.TextWidth(sTemp) > picPreview.ScaleWidth Then
                            nSplit2 = I - 1
                            Exit For

                        End If

                    Next

                    If nSplit2 = 0 Then
                        nSplit2 = UBound(sWork)

                    End If

                    For I = nSplit1 To nSplit2
                        sText2 = sText2 & IIf(I > 0, " ", "") & sWork(I)
                    Next
                    
                    'Print Text
                    With picPreview

                        If sChunk(1) = "c" Then
                            iCalc = (.ScaleWidth - .TextWidth(sText2)) / 2
                            .CurrentX = IIf(iCalc > 0, iCalc, 0)
                        ElseIf CLng(sChunk(1)) < 0 Then
                            .CurrentX = .ScaleWidth - .TextWidth(sText2) + CLng(sChunk(1))
                        Else 'If CLng(sChunk(1)) >= 0 Then
                            .CurrentX = CLng(sChunk(1))

                        End If

                        .CurrentY = CLng(sChunk(2)) + .TextHeight(sText1) + 1

                    End With
                    
                    picPreview.Print sText2

                End If
                
        End Select
        
PreviewLabel_BAD_DATA:
    Next
    
PreviewLabel_EXIT:
    picPreview.refresh
    Me.refresh
    
    Erase sSampData
    Erase sChunk
    Erase sline
    Erase sWork
    
    Me.MousePointer = vbNormal
    
End Sub

Private Function VerifyLineData(ByVal sLineData As String) As Boolean

    VerifyLineData = True
    
    Dim sChunk() As String

    Dim sErrMsg  As String
    
    sErrMsg = ""

    If LenB(sLineData) > 0 Then
        sChunk = Split(sLineData, "|")
    Else
        ReDim sChunk(0)
        sChunk(0) = ""

    End If
    
    'Verify the line data based on first character (ie line type)
    Select Case sChunk(0)

        Case "s"

            If sChunk(1) = "c" Then
                If sChunk(2) <> "I25" And sChunk(2) <> "C39" Then
                    sErrMsg = "La siguiente linea," & vbcrlf & vbcrlf & sLineData & vbcrlf & vbcrlf & "contains an invalid value for Barcode Format."

                End If

            ElseIf sChunk(1) = "o" Then

                If sChunk(2) <> "p" And sChunk(2) <> "l" Then
                    sErrMsg = "La siguiente Linea," & vbcrlf & vbcrlf & sLineData & vbcrlf & vbcrlf & "contains an invalid value for orientation."
                ElseIf UBound(sChunk) > 2 Then
                    sErrMsg = "La siguiente Linea," & vbcrlf & vbcrlf & sLineData & vbcrlf & vbcrlf & "contains extra information."

                End If

            ElseIf sChunk(1) = "s" Then

                If sChunk(2) <> "in" And sChunk(2) <> "pix" And sChunk(2) <> "cm" Then
                    sErrMsg = "La siguiente Linea," & vbcrlf & vbcrlf & sLineData & vbcrlf & vbcrlf & "contains an invalid value for scale."
                ElseIf Not (IsNumeric(sChunk(3)) And IsNumeric(sChunk(4))) Then
                    sErrMsg = "La siguiente Linea," & vbcrlf & vbcrlf & sLineData & vbcrlf & vbcrlf & "contains an invalid value for height and/or width."
                ElseIf UBound(sChunk) > 4 Then
                    sErrMsg = "La siguiente Linea," & vbcrlf & vbcrlf & sLineData & vbcrlf & vbcrlf & "contains extra information."

                End If

            ElseIf sChunk(1) = "f" Then

                If Not IsNumeric(sChunk(3)) Then
                    sErrMsg = "La siguiente Linea," & vbcrlf & vbcrlf & sLineData & vbcrlf & vbcrlf & "contains an invalid value for the font size."
                ElseIf sChunk(4) <> "true" And sChunk(4) <> "false" Then
                    sErrMsg = "La siguiente Linea," & vbcrlf & vbcrlf & sLineData & vbcrlf & vbcrlf & "contains an invalid value for the bold setting."
                ElseIf sChunk(5) <> "true" And sChunk(5) <> "false" Then
                    sErrMsg = "La siguiente linea," & vbcrlf & vbcrlf & sLineData & vbcrlf & vbcrlf & "Contiene un valor no valido para establecer la letra cursiva."
                ElseIf UBound(sChunk) > 5 Then
                    sErrMsg = "La siguiente linea," & vbcrlf & vbcrlf & sLineData & vbcrlf & vbcrlf & "Contiene informacion extra."

                End If

            Else
                sErrMsg = "La siguiente linea," & vbcrlf & vbcrlf & sLineData & vbcrlf & vbcrlf & "Contiene un tipo de configuracion invalida."

            End If

        Case "f"

            If UBound(sChunk) > 0 Then
                sErrMsg = "La siguiente linea," & vbcrlf & vbcrlf & sLineData & vbcrlf & vbcrlf & "Contiene informacion extra," & vbcrlf & "Cuando solo una f es requerido."

            End If
            
        Case "b", "t", "d"

            If Not ((sChunk(1) = "c" Or IsNumeric(sChunk(1))) And IsNumeric(sChunk(2))) Then
                sErrMsg = "La siguiente Linea," & vbcrlf & vbcrlf & sLineData & vbcrlf & vbcrlf & "Contiene un valor de posicion invalido."
            ElseIf UBound(sChunk) > 3 Then
                sErrMsg = "La siguiente Linea," & vbcrlf & vbcrlf & sLineData & vbcrlf & vbcrlf & "Contiene informacion extra."

            End If
            
        Case Else

            If LenB(Trim(sLineData)) = 0 Then
                sErrMsg = "Una Linea en blanco se encontro." & vbcrlf & "Por favor borre la linea en blanco."
            Else
                sErrMsg = "La siguiente Linea:" & vbcrlf & vbcrlf & sLineData & vbcrlf & vbcrlf & "Contiene informacion invalida."

            End If
            
    End Select
    
    Erase sChunk
    
    If LenB(sErrMsg) > 0 Then
        VerifyLineData = False
        MsgBox sErrMsg & vbcrlf & vbcrlf & "Please verify the Label definition.", vbExclamation + vbOKOnly, "Error Generating the Label"

    End If
    
End Function
