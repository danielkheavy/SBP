Attribute VB_Name = "Module24"
Option Explicit

Public Type DOCINFO

    pDocName As String
    pOutputFile As String
    pDatatype As String

End Type

Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long

Public Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long

Public Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long

Public Declare Function OpenPrinter _
               Lib "winspool.drv" _
               Alias "OpenPrinterA" (ByVal pPrinterName As String, _
                                     phPrinter As Long, _
                                     ByVal pDefault As Long) As Long

Public Declare Function StartDocPrinter _
               Lib "winspool.drv" _
               Alias "StartDocPrinterA" (ByVal hPrinter As Long, _
                                         ByVal Level As Long, _
                                         pDocInfo As DOCINFO) As Long

Public Declare Function StartPagePrinter _
               Lib "winspool.drv" (ByVal hPrinter As Long) As Long

Public Declare Function WritePrinter _
               Lib "winspool.drv" (ByVal hPrinter As Long, _
                                   pBuf As Any, _
                                   ByVal cdBuf As Long, _
                                   pcWritten As Long) As Long

Public Declare Function EndPage Lib "winspool.drv" (ByVal hPrinter As Long) As Long

Global lhPrinter As Long, lreturn As Long, lpcWritten As Long, lDoc As Long, sWrittenData As String, MyDocInfo As DOCINFO

Sub archivo_imprimir(buf As String)

    On Error GoTo cmd909090_err

    Dim mivariable As String

    Open buf For Input As #15

    While Not EOF(15)

        Line Input #15, mivariable
    Wend
    Close #15
    Exit Sub
cmd909090_err:
    MsgBox "Aviso en Archivo Imprimir " + error$, 48, "Aviso"
    Exit Sub

End Sub

Public Sub Imprimir(Texto As String, _
                    Optional ancho As Integer = 40, _
                    Optional Salto As Boolean = True)

    If Len(Texto) > ancho Then
        Texto = Left(Texto, ancho)

    End If

    If Salto Then Texto = Texto + vbcrlf
    lreturn = WritePrinter(lhPrinter, ByVal Texto, Len(Texto), lpcWritten)

End Sub

Public Sub Iniciar_Impresion(DocName As String)
    lreturn = OpenPrinter(Printer.DeviceName, lhPrinter, 0)

    '**********************************************************
    If lreturn = 0 Then
        MsgBox "The Printer Name you typed wasn't recognized."
        Exit Sub

    End If

    '***************************************************
    MyDocInfo.pDocName = DocName
    MyDocInfo.pOutputFile = vbNullString
    MyDocInfo.pDatatype = vbNullString
    lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
    Call StartPagePrinter(lhPrinter)

    'comprimido
    Rem lreturn = WritePrinter(lhPrinter, ByVal Chr$(15) + vbCrLf, Len(Chr$(15) + vbCrLf), lpcWritten)
End Sub

Public Sub Finalizar_Impresion()
    lreturn = EndPagePrinter(lhPrinter)
    lreturn = EndDocPrinter(lhPrinter)
    lreturn = ClosePrinter(lhPrinter)

End Sub
