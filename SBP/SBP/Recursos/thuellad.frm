VERSION 5.00
Begin VB.Form thuellad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Huella Digital"
   ClientHeight    =   7185
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   10065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2775
      Left            =   0
      ScaleHeight     =   2715
      ScaleWidth      =   2715
      TabIndex        =   2
      Top             =   0
      Width           =   2775
   End
   Begin VB.ListBox Status 
      Height          =   3375
      Left            =   3000
      TabIndex        =   1
      Top             =   840
      Width           =   4335
   End
   Begin VB.PictureBox HiddenPict 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   480
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label tipo 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label nombre 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   6120
      Width           =   7095
   End
   Begin VB.Label codigo 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Cursor:"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Prompt 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tocar lector huella para Leer"
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "Estado:"
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Necesita leer :"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Samples 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   4440
      Width           =   615
   End
   Begin VB.Menu flo00 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "thuellad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim WithEvents Capture As DPFPCapture
Attribute Capture.VB_VarHelpID = -1

Dim CreateFtrs         As DPFPFeatureExtraction

Dim CreateTempl        As DPFPEnrollment

Dim ConvertSample      As DPFPSampleConversion

Dim Templ              As DPFPTemplate

Option Explicit

Private Sub DrawPicture(ByVal Pict As IPictureDisp)
    ' Must use hidden PictureBox to easily resize picture.
    Set HiddenPict.Picture = Pict
    Picture1.PaintPicture HiddenPict.Picture, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, 0, 0, HiddenPict.ScaleWidth, HiddenPict.ScaleHeight, vbSrcCopy
    Picture1.Picture = Picture1.Image

End Sub

Private Sub ReportStatus(ByVal Str As String)
    ' Add string to list box.
    Status.AddItem (Str)
    ' Move list box selection down.
    Status.ListIndex = Status.NewIndex

End Sub

Private Sub Close_Click()
    ' Stop capture operation. This code is optional.
    Capture.StopCapture
    ' Unload form.
    Unload Me

End Sub

Private Sub flo00_Click()
    Capture.StopCapture
    'thuella.Hide 1
    Unload thuellad

End Sub

Private Sub Form_Load()
    ' Create capture operation.
    Set Capture = New DPFPCapture
    ' Start capture operation.
    Capture.StartCapture
    ' Create DPFPFeatureExtraction object.
    Set CreateFtrs = New DPFPFeatureExtraction
    ' Create DPFPEnrollment object.
    Set CreateTempl = New DPFPEnrollment
    ' Show number of samples needed.
    Samples.Caption = CreateTempl.FeaturesNeeded
    ' Create DPFPSampleConversion object.
    Set ConvertSample = New DPFPSampleConversion

End Sub

Private Sub Capture_OnReaderConnect(ByVal ReaderSerNum As String)
    ReportStatus ("El lector de huellas digitales se conectó.")

End Sub

Private Sub Capture_OnReaderDisconnect(ByVal ReaderSerNum As String)
    ReportStatus ("The fingerprint reader was disconnected.")

End Sub

Private Sub Capture_OnFingerTouch(ByVal ReaderSerNum As String)
    ReportStatus ("El lector de huellas digitales se desconectó.")

End Sub

Private Sub Capture_OnFingerGone(ByVal ReaderSerNum As String)
    ReportStatus ("El dedo se retira del lector de huellas digitales.")

End Sub

Private Sub Capture_OnSampleQuality(ByVal ReaderSerNum As String, _
                                    ByVal Feedback As DPFPCaptureFeedbackEnum)

    If Feedback = CaptureFeedbackGood Then
        ReportStatus ("La calidad de la muestra de la huella digital es bueno.")
    Else
        ReportStatus ("La calidad de la muestra de la huella digital es pobre.")

    End If

End Sub

Private Sub Capture_OnComplete(ByVal ReaderSerNum As String, ByVal Sample As Object)

    Dim Feedback As DPFPCaptureFeedbackEnum

    On Error GoTo cm8912_err

    ReportStatus ("La muestra de huellas dactilares fue capturado.")
    ' Draw fingerprint image.
    DrawPicture ConvertSample.ConvertToPicture(Sample)
    ' Process sample and create feature set for purpose of enrollment.
    Feedback = CreateFtrs.CreateFeatureSet(Sample, DataPurposeEnrollment)

    ' Quality of sample is not good enough to produce feature set.
    If Feedback = CaptureFeedbackGood Then
        ReportStatus ("El conjunto de características de huellas dactilares se creó.")
        Prompt.Caption = "Toque el lector de huellas digitales de nuevo con el mismo dedo."
        ' Add feature set to template.
        CreateTempl.AddFeatures CreateFtrs.FeatureSet
        ' Show number of samples needed to complete template.
        Samples.Caption = CreateTempl.FeaturesNeeded

        ' Check if template has been created.
        If CreateTempl.TemplateStatus = TemplateStatusTemplateReady Then
            thuellad.SetTemplete CreateTempl.Template
            ' Template has been created, so stop capturing samples.
            Capture.StopCapture
            SaveTemplate_Click
            Prompt.Caption = "Haga clic en Cerrar y, a continuación, haga clic en la huella digital de Verificación."
            MsgBox "La plantilla de huella digital fue creado."

        End If

    End If

    Exit Sub
cm8912_err:
    MsgBox "Aviso en captura " + error$, 48, "Aviso"
    Exit Sub

End Sub

Public Sub SetTemplete(ByVal Template As Object)
    Set Templ = Template

End Sub

Private Sub SaveTemplate_Click()

    On Error GoTo cmd9090_err

    Dim blob() As Byte

    ' First verify that template is not empty.
    'If Templ Is Nothing Then
    ' MsgBox "You must create a fingerprint template before you can save it."
    ' Exit Sub
    'End If
    'CommonDialog1.Filter = "Fingerprint Template File|*.fpt"
    ' Set dialog box so an error occurs if dialog box is cancelled.
    'CommonDialog1.CancelError = True
    'On Error Resume Next
    ' Show Save As dialog box.
    'CommonDialog1.ShowSave
    'If Err Then
    ' This code runs if the dialog box was cancelled.
    '   Exit Sub
    'End If
    ' Export template to binary data.
    blob = Templ.Serialize
    ' Save binary data to file.
    'MsgBox globalpath & "001d\06\huella\" & Trim(codigo) & ".fpt"
    Open globalpath & "\001d\06\huella\" & Trim(tipo) & "\" & Trim(codigo) & ".fpt" For Binary As #119
    Put #119, , blob
    Close #119
    Exit Sub
cmd9090_err:
    MsgBox "Aviso en grabar Huella " + error$, 48, "Aviso"
    Exit Sub

End Sub

