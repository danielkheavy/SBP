VERSION 5.00
Begin VB.Form thuellat 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test de Huella"
   ClientHeight    =   6030
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   7860
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2775
      Left            =   0
      ScaleHeight     =   2715
      ScaleWidth      =   2715
      TabIndex        =   11
      Top             =   0
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   8280
      TabIndex        =   9
      Top             =   1560
      Width           =   3615
   End
   Begin VB.ListBox Status 
      Height          =   1815
      Left            =   3000
      TabIndex        =   1
      Top             =   840
      Width           =   4575
   End
   Begin VB.PictureBox HiddenPict 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   840
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label tipo 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Label codigo 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label nombre 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5040
      Width           =   7335
   End
   Begin VB.Label Label2 
      Caption         =   "Estado:"
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Prompt 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Toque el lector de huellas digitales."
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Prompt:"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Tasa de Falsa Aceptación:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label FAR 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Menu lfdo92 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "thuellat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim WithEvents Capture As DPFPCapture
Attribute Capture.VB_VarHelpID = -1

Dim CreateFtrs         As DPFPFeatureExtraction

Dim Verify             As DPFPVerification

Dim ConvertSample      As DPFPSampleConversion

Dim Templ              As DPFPTemplate

Option Explicit

Private Sub ReportStatus(ByVal Str As String)
    ' Add string to list box.
    Status.AddItem (Str)
    ' Move list box selection down.
    Status.ListIndex = Status.NewIndex

End Sub

Private Sub DrawPicture(ByVal Pict As IPictureDisp)
    ' Must use hidden PictureBox to easily resize picture.
    Set HiddenPict.Picture = Pict
    Picture1.PaintPicture HiddenPict.Picture, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, 0, 0, HiddenPict.ScaleWidth, HiddenPict.ScaleHeight, vbSrcCopy
    Picture1.Picture = Picture1.Image

End Sub

Private Sub Close_Click()

End Sub

Private Sub Form_Activate()
    'If Len(Trim(codigo)) > 0 Then
    '   ReadTemplate_Click
    'End If
    codigohuella = ""
    carga_directorio

End Sub

Function Extraer(path As String, Caracter As String) As String

    Dim ret               As String

    Dim posicionextension As Integer

    posicionextension = InStrRev(path, Caracter)

    If posicionextension <> 0 Then
        ret = Trim(Mid$(path, 1, posicionextension - 1))

    End If

    ' -- Retorna el valor
    Extraer = ret

End Function

Sub buscar_directorio()

    Dim I As Integer

    For I = 1 To List1.ListCount
        codigo = List1.List(I)
        ReadTemplate_Click
    Next I

End Sub

Private Sub Form_Load()

    On Error GoTo cmd89000_err

    ' Create capture operation.
    Set Capture = New DPFPCapture
    ' Start capture operation.
    Capture.StartCapture
    ' Create DPFPFeatureExtraction object.
    Set CreateFtrs = New DPFPFeatureExtraction
    ' Create DPFPVerification object.
    Set Verify = New DPFPVerification
    ' Create DPFPSampleConversion object.
    Set ConvertSample = New DPFPSampleConversion
    Exit Sub
cmd89000_err:
    MsgBox "Aviso en Load " + error$, 48, "Aviso"
    Exit Sub
 
End Sub

Private Sub Capture_OnReaderConnect(ByVal ReaderSerNum As String)
    ReportStatus ("Lector esta conectado.")

End Sub

Private Sub Capture_OnReaderDisconnect(ByVal ReaderSerNum As String)
    ReportStatus ("Lector Desconectado")

End Sub

Private Sub Capture_OnFingerTouch(ByVal ReaderSerNum As String)
    ReportStatus ("Lector fue Tocado")

End Sub

Private Sub Capture_OnFingerGone(ByVal ReaderSerNum As String)
    ReportStatus ("Dedo retirado del Lector")

End Sub

Private Sub Capture_OnSampleQuality(ByVal ReaderSerNum As String, _
                                    ByVal Feedback As DPFPCaptureFeedbackEnum)

    If Feedback = CaptureFeedbackGood Then
        ReportStatus ("La calidad de la muestra de la huella digital es buena.")
    Else
        ReportStatus ("La calidad de la muestra de la huella digital es pobre.")

    End If

End Sub

Private Sub Capture_OnComplete(ByVal ReaderSerNum As String, ByVal Sample As Object)

    Dim Feedback As DPFPCaptureFeedbackEnum

    Dim Res      As DPFPVerificationResult

    Dim Templ    As Object

    ReportStatus ("La huella digital fue capturado.")
    ' Draw fingerprint image.
    'MsgBox "abc"
    DrawPicture ConvertSample.ConvertToPicture(Sample)
    ' Process sample and create feature set for purpose of verification.
    Feedback = CreateFtrs.CreateFeatureSet(Sample, DataPurposeVerification)
    validaciones
    Exit Sub

    ' Quality of sample is not good enough to produce feature set.
    If Feedback = CaptureFeedbackGood Then
        Prompt.Caption = "Toque el lector de huellas digitales con un dedo diferente."
        Set Templ = thuellat.GetTemplate

        If Templ Is Nothing Then
            MsgBox "Debe crear una plantilla de huellas dactilares antes de poder realizar la verificación."
        Else
            ' Compare feature set with template.
            Set Res = Verify.Verify(CreateFtrs.FeatureSet, Templ)
            ' Show results of comparison.
            FAR.Caption = Res.FARAchieved
    
            If Res.Verified = True Then
                ReportStatus ("La huella dactilar se verificó.")
                'MsgBox "pertenece a " & codigo
                codigohuella = Trim(codigo)
            Else
                ReportStatus ("La huella digital no se verificó.")

            End If

        End If

    Else
        ReportStatus ("La calidad del conjunto de características es pobre.")

    End If

End Sub

Private Sub ReadTemplate_Click()

    On Error GoTo cmd90123_err

    Dim blob() As Byte

    'CommonDialog1.Filter = "Fingerprint Template File|*.fpt"
    ' Set dialog box so an error occurs if dialog box is cancelled.
    'CommonDialog1.CancelError = True
    'On Error Resume Next
    ' Show Open dialog box.
    'CommonDialog1.ShowOpen
    'If Err Then
    ' This code runs if dialog box was cancelled.
    '   Exit Sub
    'End If
    ' Read binary data from file.
    'MsgBox globalpath & "001d\06\huella\" & Trim(codigo) & ".fpt"
    Open globalpath & "\001d\06\huella\" & Trim(tipo) & "\" & Trim(codigo) & ".fpt" For Binary As #1
    ReDim blob(LOF(1))
    Get #1, , blob()
    Close #1

    ' Template can be empty, it must be created first.
    If Templ Is Nothing Then Set Templ = New DPFPTemplate
    ' Import binary data to template.
    Templ.Deserialize blob
    Exit Sub
cmd90123_err:
    MsgBox "No se puede leer " + error$, 48, "Aviso"
    Exit Sub

End Sub

Function GetTemplate() As Object

    ' Template can be empty. If so, then returns Nothing.
    If Templ Is Nothing Then
    Else
        Set GetTemplate = Templ

    End If

End Function

Sub lfdo92_Click()
    ' Stop capture operation. This code is optional.
    Capture.StopCapture
    ' Unload form.
    Unload Me

End Sub

Sub carga_directorio()

    On Error GoTo cmd90678_err

    Dim directory As String

    Dim R         As Integer

    Dim f

    directory = globalpath & "\001d\06\huella\" & Trim(tipo) & "\"
    'MsgBox directory
    List1.Clear
    R = 1
    f = Dir(directory, 16)

    Do While f <> ""
        R = R + 1
        f = Dir()

        If f <> ".." Then
            List1.AddItem Extraer("" & f, ".")
   
        End If    'Esto añade el archivo actual en el ListBox

    Loop
    Exit Sub
cmd90678_err:
    MsgBox "Aviso en cargar directorio " + error$, 48, "Aviso"
    Exit Sub

    'List1.AddItem 'Esto añade el directorio como último elemento
End Sub

Function valida_carga()

    On Error GoTo cmd9012345_err

    Dim Feedback As DPFPCaptureFeedbackEnum

    Dim Res      As DPFPVerificationResult

    Dim Templ    As Object

    If Feedback = CaptureFeedbackGood Then
        Prompt.Caption = "Toque el lector de huellas digitales con un dedo diferente."
        Set Templ = thuellat.GetTemplate

        If Templ Is Nothing Then
            MsgBox "Debe crear una plantilla de huellas dactilares antes de poder realizar la verificación."
        Else
            ' Compare feature set with template.
            Set Res = Verify.Verify(CreateFtrs.FeatureSet, Templ)
            ' Show results of comparison.
            FAR.Caption = Res.FARAchieved
    
            If Res.Verified = True Then
                ReportStatus ("La huella dactilar se verificó.")
                'MsgBox "pertenece a " & codigo
                valida_carga = 1
            Else
                ReportStatus ("La huella digital no se verificó.")

            End If

        End If

    Else
        ReportStatus ("La calidad del conjunto de características es pobre.")

    End If

    Exit Function
cmd9012345_err:
    MsgBox "Aviso en valida carga " + error$, 48, "Aviso"
    Exit Function

End Function

Sub validaciones()

    Dim I As Integer

    codigohuella = ""

    For I = 0 To List1.ListCount - 1

        If Len(Trim("" & List1.List(I))) > 0 Then
            codigo = Trim(List1.List(I))
            ReadTemplate_Click

            If valida_carga() = 1 Then
                codigohuella = Trim(codigo)
                lfdo92_Click
                Exit Sub

            End If

        End If

    Next I

End Sub

