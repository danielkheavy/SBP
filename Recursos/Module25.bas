Attribute VB_Name = "Module25"
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrn As Long, pDefault As Any) As Long
Private Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hprn As Long, ByVal Level As Long, pDocInfo As DOC_INFO_1) As Long
Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal hprn As Long) As Long
Private Declare Function WritePrinter Lib "winspool.drv" (ByVal hprn As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long
Private Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hprn As Long) As Long
Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hprn As Long) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hprn As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Type DOC_INFO_1
   pDocName As String
   pOutputFile As String
   pDatatype As String
End Type

Dim bien As Long
Public Sub Iniciar_Spool(PrnName As String, hprn As Long)
'Abrir impresora, donde PRnName es el DeviceName de la impresora.Se devuelve hprn como identificador
  Dim di As DOC_INFO_1
  di.pDocName = "Cadena"
  di.pOutputFile = vbNullString
  di.pDatatype = "RAW"
  Call OpenPrinter(PrnName, hprn, ByVal 0&)
  Call StartDocPrinter(hprn, 1, di)
  Call StartPagePrinter(hprn)
End Sub
Public Sub Fin_Spool(ByVal hprn As Long)
'Cerrar impresora
  Call EndPagePrinter(hprn)
  Call EndDocPrinter(hprn)
  Call ClosePrinter(hprn)
End Sub



Public Sub imprimir_Spool(ByVal strCadena As String, ByVal hprn As Long, salto As Integer)
'Imprimir cadena


  Dim buffer() As Byte
  Dim Written As Long
  Dim i As Long, j As Long
  Const BufSize As Long = &H50 '80 CARACTERES

  Dim BytesStr As Long
  Dim acopiar$
  Dim csalto$
  Dim cretorno$
  Dim veces As Integer
  Dim asciicadena As String

  asciicadena = strCadena 'cadena_ansi_a_ascii(strCadena)
  cretorno = Chr(10) 'Salto linea
  csalto = Chr(13) 'Retorno carro
  veces = 1
  If salto Then veces = 3
  Do While veces > 0
    If salto Then
      If veces = 3 Then 'Cadena
        acopiar = asciicadena
      Else
        If veces = 2 Then 'Salto
          acopiar = csalto
        Else 'Retorno
          acopiar = cretorno
        End If
      End If
    Else 'No salto
      acopiar = asciicadena
    End If
    BytesStr = ((Len(acopiar)) * 2) - 1
    ReDim buffer(1 To BufSize) As Byte
    For i = 1 To UBound(buffer)
      buffer(i) = &H20 'Rellenar a blancos el buffer
    Next i
    buffer() = acopiar$ 'Llenar el buffer
    Call WritePrinter(hprn, buffer(0), BytesStr, Written) 'Imprimir
    veces = veces - 1
  Loop

End Sub
Sub prueba_imprime(xbuf As String)
Dim himp As Long
Dim txtTheLine As String
Call Iniciar_Spool("cajax", himp)
Open xbuf For Input As 1
Do While Not EOF(1)
Line Input #1, txtTheLine '& VarTexto & vbCrLf   'Read a line
'Printer.Print txtTheLine & vbcrlf     'send to the default printer.
Call imprimir_Spool(txtTheLine & vbcrlf, himp, True)
Loop
Close #1
Call Fin_Spool(himp)
End Sub
