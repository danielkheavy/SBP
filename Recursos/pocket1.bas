Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Sub RtlMoveMemory Lib "kernel32" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
'Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)
Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Public Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)
Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Global rrlocal11 As String
Global rrtipo As String
Global rrserie As String
Global rrnumero As String
Global anticipoo As String
Global sw_consulta As Integer

Global dbclie As New ADODB.Recordset  'ojo es general

Global flag_comanda As String
Global serial_number As String
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40
Global tipodescuento As String
Global valordescuento As Double
Global glocal As String
Global txtotlare As Double
Global mysnap As Snapshot
Global tipoletra As String
Global opcion5 As Integer
Global orionv4 As String
Global suma1 As Double
Global suma2 As Double
Global suma3 As Double
Global suma4 As Double

Global suma5 As Double
Global suma6 As Double
Global suma7 As Double
Global suma8 As Double
Global suma9 As Double

Global ssuma1 As Double
Global ssuma2 As Double
Global ssuma3 As Double
Global ssuma4 As Double

Global ssuma5 As Double
Global ssuma6 As Double
Global ssuma7 As Double
Global ssuma8 As Double
Global ssuma9 As Double
Global amsw As Integer

Global mensaje_print As String
Global xxxsoles As String
'Global unidadx As String
Global opcion1 As String
Global opcion2 As String
Global opcion3 As String
Global signopeso As String
Global mensaje_bloqueo As String
Global globaldir As String
Global globaldat As String
Global gglobaldir As String
Global gglobaldat As String
Global globalpath As String
Global globalweb As String
Global globalemp As String
Global globalocal As String
Global empresapos As String
Global mytablexxx As Table
Global mydbxxx As Database



Global globalpri As String
Global globalcont As String
Global gusuario As String
Global FileName As String
Global ncanal As Integer
Global mydb11 As Database
'Global mytable11 As Table
Global mydbxglo As Database
Global mydbzglo As Database 'contable

Global flag_contando As Integer   'flag de nro item en el formato factura
Global dbserial As String

Global xarchivo As String
Global xarchivo1  As String
Global dia As String

Global tipo_servicio As String
Global cgusuario As String
Global usuariopos As String
Global contlin As Double
Global contpag As Double
Global ticketera_cajon As Integer

Global dgusuario As String
Global dgusuariog As String
Global fgusuario As String
Global fpusuario As String
Global fpusuarior As String
Global gocabeza As String
Global godetalle As String
Global gofpago As String

      Global dbbase As String
      Global dbca As String
      Global dbing As String
      Global dbde As String
      Global dbfp As String
      Global dbtalla As String
      Global xnpuerto As String
      Global xnpuerto1 As String
      Global flag_clave1 As Integer
Sub selecciona_impresoras(Nombre_Impresora As String)
On Error GoTo cm8911_err
Dim Prt As Printer
    ' Establece la impresora que se utilizar¨¢ para imprimir
    For Each Prt In Printers
        If Prt.DeviceName = Nombre_Impresora Then
            Set Printer = Prt
        End If
    Next
    'Printet = Nothing
    Exit Sub
cm8911_err:
    MsgBox "No se puede configurar Impresora " + error$, 48, "Aviso"
    Exit Sub


End Sub
      
Function formateaa(buf As String, longitud As Integer, sw As Integer, sw1 As Integer)
Dim xbuf As String
Dim buf1 As String
Dim sdx As Integer
On Error GoTo cmd200_err
buf1 = buf
sdx = longitud - Len(buf)
If sdx > 0 Then
   If sw1 = 0 Then
      buf1 = buf & Space$(sdx)
   End If
   If sw1 = 1 Then
      buf1 = Space$(sdx) & buf
   End If
End If
xbuf = Mid$(buf1, 1, longitud)
If sw = 0 Then
   'Printer.Print xbuf;
   Print #1, xbuf;
End If
If sw = 1 Then
   'Printer.Print xbuf,
   Print #1, xbuf,
End If
If sw = 2 Then
   'Printer.Print xbuf
   Print #1, xbuf
End If
'Close #ncanal
Exit Function
cmd200_err:
'MsgBox Error$
MsgBox "Mensaje, Error en formateaa " & error$, 24, "Aviso"
Exit Function
End Function
Public Function IsValidIPAddress(ByVal strIPAddress As String) As Boolean
    On Error GoTo Handler
    Dim varAddress As Variant, n As Long, lCount As Long
    varAddress = Split(strIPAddress, ".", , vbTextCompare)
    '//
    If IsArray(varAddress) Then
        For n = LBound(varAddress) To UBound(varAddress)
            lCount = lCount + 1
            varAddress(n) = CByte(varAddress(n))
        Next
        '//
        IsValidIPAddress = (lCount = 4)
    End If
    '//
Handler:
End Function
Function serie_disco_duro() As String
Dim unidad As String
Dim buf As String
Dim xbuf As String
Dim ybuf As String

On Error GoTo cmd9011_err
Dim cad1 As String * 256
    Dim cad2 As String * 256
    Dim numSerie As Long
    Dim longitud As Long
    Dim flag As Long
    Dim found As Integer
    '
    
    'anterior fueron estos dos
    'unidad = "C:\"
    'Call GetVolumeInformation(unidad, cad1, 256, numSerie, longitud, flag, cad2, 256)
    'buf = numSerie
    'fin a de anterior
    
    'ahora haciendo con la serie disco duro
    'serie_disco_duro = placa_madre()
    'MsgBox "Numero de Serie de la unidad " & unidad & " = " & numSerie
    xbuf = ""
    'ybuf = "" & pocket.vservidor
    'If Len(menup.vservidor) > 0 Then
       'If GetRemoteMACAddress(ybuf, xbuf, "") Then
       ' xbuf = xbuf
       ' Else
       ' xbuf = ""
       'End If
       'MsgBox xbuf
       
    '   End If
    'End If
    'MsgBox xbuf
    
    unidad = "C:\"
    Call GetVolumeInformation(unidad, cad1, 256, numSerie, longitud, flag, cad2, 256)
    buf = xbuf + "" & numSerie
    serie_disco_duro = buf
    'MsgBox "paso"
    Exit Function
cmd9011_err:
    MsgBox "Aviso en shd" + error$, 48, "Aviso"
    Exit Function
End Function

Sub borrar_archivo(buf As String)
On Error GoTo cmd34_err
If Dir$(buf) = "" Then
Exit Sub
End If
Kill buf
Exit Sub
cmd34_err:
Exit Sub
End Sub
Public Function existe_archivo(sArchivo As String) As Integer
On Error GoTo cmd6823_err
    existe_archivo = Len(Dir$(sArchivo))
    Exit Function
cmd6823_err:
Exit Function
End Function
Function bFileExists(FileName As String) As Integer
    Dim i As Integer
    On Error Resume Next
    i = Len(Dir$(FileName))
    If Err Or i = 0 Then
        bFileExists = False
    Else
        bFileExists = True
    End If
End Function
Function existearchivo(FileName As String) As Integer
    Dim i As Integer
    On Error Resume Next
    i = Len(Dir$(FileName))
    If Err Or i = 0 Then
        existearchivo = 0
    Else
        existearchivo = 1
    End If
End Function

Sub copiar_archivo(buf As String, antdir As String)
    On Error GoTo cmd56_err
    Dim buf1 As String
    Screen.MousePointer = 11
    FileCopy antdir, buf   ' Copy file to destination.
    Screen.MousePointer = 1
    Exit Sub
cmd56_err:
    MsgBox "ERROR AL COPIAR...", 24, "AVISO"
    Screen.MousePointer = 1
    Exit Sub
End Sub




Function copiar_temporal()
On Error GoTo cmd23_err
FileCopy globaldat & "\tfactura.dbf", globaldat & "\" & "_c" & gusuario & ".dbf"
FileCopy globaldat & "\tfactura.cdx", globaldat & "\" & "_c" & gusuario & ".cdx"
FileCopy globaldat & "\tdetalle.dbf", globaldat & "\" & "_d" & gusuario & ".dbf"
FileCopy globaldat & "\tdetalle.cdx", globaldat & "\" & "_d" & gusuario & ".cdx"
copiar_temporal = 1
Exit Function
cmd23_err:
Exit Function
End Function
Function copiar_deliveri(buf As String)
On Error GoTo cmd223_err
FileCopy globaldat & "\tdetalle.dbf", globaldat & "\" & "_z" & buf & ".dbf"
FileCopy globaldat & "\tdetalle.cdx", globaldat & "\" & "_z" & buf & ".cdx"
copiar_deliveri = 1
Exit Function
cmd223_err:
MsgBox "Error al Copiar temporal Detalle " + error$, 48, "Aviso"
Exit Function

End Function
Function copiar_temporalpe()
On Error GoTo cmd23_err
FileCopy globaldat & "\tmppedir.dbf", globaldat & "\" & "_t" & gusuario & ".dbf"
FileCopy globaldat & "\tmppedir.cdx", globaldat & "\" & "_t" & gusuario & ".cdx"
copiar_temporalpe = 1
Exit Function
cmd23_err:
Exit Function
End Function

Function copiar_tcxcre()
On Error GoTo cmd22333_err
'MsgBox globaldat & "\tcxcre.dbf"
'MsgBox globaldat & "\" & "_b" & gusuario & ".dbf"
FileCopy globaldat & "\tcxcre.dbf", globaldat & "\" & "_b" & gusuario & ".dbf"
FileCopy globaldat & "\tcxcre.cdx", globaldat & "\" & "_b" & gusuario & ".cdx"
copiar_tcxcre = 1
Exit Function
cmd22333_err:
MsgBox error$
Exit Function
End Function

Function copiar_recibos()
On Error GoTo cmd223_err
FileCopy globaldat & "\tmpcta.dbf", globaldat & "\" & "_r" & gusuario & ".dbf"
FileCopy globaldat & "\tmpcta.cdx", globaldat & "\" & "_r" & gusuario & ".cdx"
copiar_recibos = 1
Exit Function
cmd223_err:
Exit Function

End Function
Function copiar_tmpfpago(buf As String)
On Error GoTo cmd34_err
FileCopy globaldat & "\tmpfpago.dbf", globaldat & "\" & buf & ".dbf"
'FileCopy globaldat & "\tmpfpago.cdx", globaldat & "\" & "_c" & fpusuario & ".cdx"
'MsgBox globaldat & "\" & "_f" & buf & ".dbf"
copiar_tmpfpago = 1
Exit Function
cmd34_err:
MsgBox " Error al Copiar FormaPago ", 48, "Aviso"
Exit Function
End Function
Function copiar_tmpfpagoR()
On Error GoTo cmd34_err
FileCopy globaldat & "\tmpfpago.dbf", globaldat & "\" & "_l" & gusuario & ".dbf"
'FileCopy globaldat & "\tmpfpago.cdx", globaldat & "\" & "_c" & fpusuario & ".cdx"
copiar_tmpfpagoR = 1
Exit Function
cmd34_err:
Exit Function
End Function


Function copiar_tmprodu()
On Error GoTo cmd38_err
FileCopy globaldat & "\tmprodu.dbf", globaldat & "\" & "_p" & gusuario & ".dbf"
FileCopy globaldat & "\tmprodu.cdx", globaldat & "\" & "_p" & gusuario & ".cdx"
copiar_tmprodu = 1
Exit Function
cmd38_err:
Exit Function
End Function

Function proceso_formatos(archivo_formato As String, mytablex As ADODB.Recordset, ubicacioni As String, ubicacionf As String, basedatos As String, indice As String, bxlocal As String, tipo As String, bxserie As String, numero As String, ascopia As String, contando As Integer)
On Error GoTo cmd56789_err
Dim linea$
Dim buff$
Dim campo As String
Dim j As Integer
Dim sw As Integer
Dim posicioni As Long
Dim posicionf As Long
Dim tlinea As String
Dim valor As String
Dim found As Integer
Dim nombrearch As String
Dim nombrearch1 As String
Dim posicionb As Long
Dim variable As String
Dim sw1 As Integer
Dim bufx As String
Dim xxsw As Integer
Dim alibaba As Integer
    cerrar_archivo
    nombrearch = globaldir & "\temporal\" & gusuario & ".txt"
    nombrearch1 = globaldir & "\formatos\" & archivo_formato
    posicionb = 1
    sw1 = 0
    ncanal = 2
    Open nombrearch For Append As #1
    Open nombrearch1 For Input As #2
Iniciado:
    xxsw = 0
    Do
       alibaba = 0
       If EOF(2) Then Exit Do
          'linea = Space$(350)
          On Error GoTo error_lectura
          Line Input #2, buff
          On Error GoTo 0
          linea = Mid$(buff, 1, Len(buff))
          If Mid$(linea, 1, 1) = ubicacioni Then
             sw1 = 1
          End If
          If Mid$(linea, 1, 1) = ubicacionf Then
             sw1 = 0
             GoTo Iniciado
          End If
          '-------------------------
          If sw1 = 1 Then  'si es cabecera
             sw = 0
             posicioni = 0
             posicionf = 0
             valor = ""
             For j = 1 To Len(linea)
                 If Mid$(linea, j, 1) = ubicacionf Then
                    sw1 = 0
                    If Mid$(campo, 1, 6) = "RECETA" Or Mid$(campo, 1, 8) = "SERIALES" Or Mid$(campo, 1, 6) = "AGRUPA" Then
                       'MsgBox "Hola"
                       GoTo Iniciado
                    End If
                    found = formateaa("", 1, 2, 0)
                    GoTo Iniciado
                 End If
                 If sw = 0 And Mid$(linea, j, 1) <> "[" And Mid$(linea, j, 1) <> "]" And Mid$(linea, j, 1) <> "{" And Mid$(linea, j, 1) <> "}" And Mid$(linea, j, 1) <> "/" And Mid$(linea, j, 1) <> "\" And Mid$(linea, j, 1) <> "<" And Mid$(linea, j, 1) <> ">" And Mid$(linea, j, 1) <> "^" And Mid$(linea, j, 1) <> "&" And Mid$(linea, j, 1) <> "$" And Mid$(linea, j, 1) <> "?" Then
                    variable = Mid$(linea, j, 1)
                    If variable <> "@" And variable <> "+" Then
                       found = formateaa(variable, 1, 0, 0)
                    End If
                 End If
                 'If Mid$(linea, j, 1) = "@" Then  'negrita
                    'bufx = Chr$(&H1B) + "!" + Chr$(8)
                    'found = formateaa(bufx, Len(bufx), 0, 0)
                 'End If
                 'If Mid$(linea, j, 1) = "+" Then  'negrita
                    'bufx = Chr$(27) + "@"
                    'found = formateaa(bufx, Len(bufx), 0, 0)
                    xxsw = 1
                 'End If
                 If Mid$(linea, j, 1) = "[" Then
                    sw = 1
                    posicioni = j + 1
                 End If
                 If sw = 1 And Mid$(linea, j, 1) = "]" Then
                    posicionf = j - 1
                    campo = Mid$(linea, posicioni, posicionf - posicioni + 1)
                    alibaba = 0
                    valor = busca_campo1(basedatos, mytablex, campo, bxlocal, bxserie, numero, contando, alibaba, ascopia)
                    sw = 0
                    posicioni = 0
                    posicionf = 0
                    If alibaba = 1 Then
                       'GoTo paseporaqui
                    End If
                 End If
             Next j
             found = formateaa("", 1, 2, 0)
paseporaqui:
          End If
          '-------------------------
    Loop
comienzar:
    Close #2
    Close #1
    cerrar_archivo
Exit Function
cmd56789_err:
    MsgBox "xx.Existe Un error en Proceso Formatos " & error$, 24, "Aviso"
    cerrar_archivo
    Exit Function
error_lectura:
    MsgBox "Error en Proceso_formatos", 24, "Aviso"
    cerrar_archivo
    Exit Function
    
End Function
Function numero_diasMes()
Dim month_number As Integer
Dim year_number As Integer

'month_number = Month(txtMonth.Text)
'year_number = Year(txtMonth.Text)
'MsgBox "Days: " & Format$(Day(DateSerial(year_number, month_number + 1, 0)))
End Function
Function busca_campo1(tablabasedatos As String, mytablex As ADODB.Recordset, campo As String, bxlocal As String, bxserie As String, tablas As String, contando As Integer, alibaba As Integer, ascopia As String) As String
Dim knik1 As String
Dim knik2 As String
Dim knik11 As String
Dim knik22 As String
Dim amigohs As String
Dim CAMPO1 As String
Dim CAMPO2 As String
Dim campo3 As String
Dim campo4 As Integer
Dim found As Integer
Dim sdx As Double
Dim campoz As String
Dim campoy As String
Dim ponemoneda As String
Dim buf As String
Dim sdx1 As Double
Dim j As Integer
Dim mytabley As New ADODB.Recordset
'Dim mydby As Database
Dim ddd As String
Dim mmm As String
Dim yyy As String
Dim bufx As String
    Dim bm
    Dim ik
On Error GoTo cmd9876_err
campo4 = 0
buf = campo
campoz = ""
campoy = ""
If InStr(campo, ">") > 0 Then  'para tomar de otra base de datos
   j = InStr(campo, ">")
   campoz = Mid$(campo, 1, j - 1)
   campoy = Mid$(campo, j + 1, Len(campo) - (j))
   '--------------------------------------------
   If campoz = "PRODUCTO" Then
       mytabley.Open "SELECT * FROM producto where  producto='" & "" & mytablex.Fields("producto") & "'", cn, adOpenDynamic, adLockOptimistic
       If mytabley.RecordCount > 0 Then 'si existe
          '----------------------------
          found = extraer_campos(campoy, CAMPO1, CAMPO2, campo3, campo4, ",")
          buf = "" & mytabley.Fields(CAMPO1)
          found = formateaa(buf, Val(campo3), 0, 0)
          '----------------------------
       End If
       mytabley.Close
       Exit Function
   End If
   
   If campoz = "TRANSPOR" Then
       mytabley.Open "SELECT * FROM transpor where  codigo='" & "" & mytablex.Fields("transporte") & "'", cn, adOpenDynamic, adLockOptimistic
       If mytabley.RecordCount > 0 Then 'si existe
          '----------------------------
          found = extraer_campos(campoy, CAMPO1, CAMPO2, campo3, campo4, ",")
          buf = "" & mytabley.Fields(CAMPO1)
          found = formateaa(buf, Val(campo3), 0, 0)
          '----------------------------
       End If
       mytabley.Close
       '
       Exit Function
   End If
   If campoz = "CLIENTES" Then
       mytabley.Open "SELECT * FROM clientes where  codigo='" & "" & mytablex.Fields("codigo") & "'", cn, adOpenDynamic, adLockOptimistic
       If mytabley.RecordCount > 0 Then 'si existe
       
          
          '----------------------------
          found = extraer_campos(campoy, CAMPO1, CAMPO2, campo3, campo4, ",")
          buf = "" & mytabley.Fields(CAMPO1)
          found = formateaa(buf, Val(campo3), 0, 0)
          '----------------------------
       End If
       mytabley.Close
       '
       Exit Function
   End If
   If campoz = "LOCALES" Then
       'Set mydby = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
       mytabley.Open "SELECT * FROM locales where  codigo='" & "" & mytablex.Fields("codigo") & "'", cn, adOpenDynamic, adLockOptimistic
       If mytabley.RecordCount > 0 Then 'si existe
          '----------------------------
          found = extraer_campos(campoy, CAMPO1, CAMPO2, campo3, campo4, ",")
          buf = "" & mytabley.Fields(CAMPO1)
          found = formateaa(buf, Val(campo3), 0, 0)
          '----------------------------
       End If
       mytabley.Close
       '
       Exit Function
   End If
   If campoz = "BODEGA" Then
       'Set mydby = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
           mytabley.Open "SELECT * FROM bodega where  codigo='" & "" & mytablex.Fields("bodega") & "'", cn, adOpenDynamic, adLockOptimistic
       If mytabley.RecordCount > 0 Then 'si existe
          '----------------------------
          found = extraer_campos(campoy, CAMPO1, CAMPO2, campo3, campo4, ",")
          buf = "" & mytabley.Fields(CAMPO1)
          found = formateaa(buf, Val(campo3), 0, 0)
          '----------------------------
       End If
       mytabley.Close
       '
       Exit Function
   End If
   If campoz = "BODEGA1" Then
       'Set mydby = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
       mytabley.Open "SELECT * FROM bodega where  bodega='" & "" & mytablex.Fields("bodegaf") & "'", cn, adOpenDynamic, adLockOptimistic
       If mytabley.RecordCount > 0 Then 'si existe
          '----------------------------
          found = extraer_campos(campoy, CAMPO1, CAMPO2, campo3, campo4, ",")
          buf = "" & mytabley.Fields(CAMPO1)
          found = formateaa(buf, Val(campo3), 0, 0)
          '----------------------------
       End If
       mytabley.Close
       '
       Exit Function
   End If
   If campoz = "VENDEDOR" Then
       'Set mydby = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
       mytabley.Open "SELECT * FROM vendedor where  codigo='" & "" & mytablex.Fields("vendedor") & "'", cn, adOpenDynamic, adLockOptimistic
       If mytabley.RecordCount > 0 Then 'si existe
       
       
          '----------------------------
          found = extraer_campos(campoy, CAMPO1, CAMPO2, campo3, campo4, ",")
          buf = "" & mytabley.Fields(CAMPO1)
          found = formateaa(buf, Val(campo3), 0, 0)
          '----------------------------
       End If
       mytabley.Close
       '
       Exit Function
   End If
     If campoz = "TIPO" Then
       'Set mydby = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
       mytabley.Open "SELECT * FROM tipo where  tipo='" & "" & mytablex.Fields("tipo") & "'", cn, adOpenDynamic, adLockOptimistic
       If mytabley.RecordCount > 0 Then 'si existe
          '----------------------------
          found = extraer_campos(campoy, CAMPO1, CAMPO2, campo3, campo4, ",")
          buf = "" & mytabley.Fields(CAMPO1)
          found = formateaa(buf, Val(campo3), 0, 0)
          '----------------------------
       End If
       mytabley.Close
       '
       Exit Function
   End If
   If campoz = "FPAGOV" Or campoz = "FPDIARIO" Then
       'Set mydby = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
        mytabley.Open "SELECT * FROM " & campoz & " where  local='" & "" & mytablex.Fields("local") & "' and tipo='" & "" & mytablex.Fields("tipo") & "' and serie='" & "" & mytablex.Fields("serie") & "' and numero=" & "" & mytablex.Fields("numero") & "'", cn, adOpenDynamic, adLockOptimistic
       If mytabley.RecordCount > 0 Then 'si existe
          '----------------------------
          found = extraer_campos(campoy, CAMPO1, CAMPO2, campo3, campo4, ",")
          buf = "" & mytabley.Fields(CAMPO1)
          found = formateaa(buf, Val(campo3), 0, 0)
          '----------------------------
       End If
       mytabley.Close
       '
       Exit Function
   End If

   If campoz = "PROVEEDO" Then
       mytabley.Open "SELECT * FROM proveedo where  codigo='" & "" & mytablex.Fields("codigo") & "'", cn, adOpenDynamic, adLockOptimistic
       If mytabley.RecordCount > 0 Then 'si existe
       'Set mydby = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
          '----------------------------
          found = extraer_campos(campoy, CAMPO1, CAMPO2, campo3, campo4, ",")
          buf = "" & mytabley.Fields(CAMPO1)
          found = formateaa(buf, Val(campo3), 0, 0)
          '----------------------------
       End If
       mytabley.Close
       '
       Exit Function
   End If
   If campoz = "FPAGO" Then
       'Set mydby = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
       mytabley.Open "SELECT * FROM fpago where  fpago='" & "" & mytablex.Fields("fpago") & "'", cn, adOpenDynamic, adLockOptimistic
       If mytabley.RecordCount > 0 Then 'si existe
          '----------------------------
          found = extraer_campos(campoy, CAMPO1, CAMPO2, campo3, campo4, ",")
          buf = "" & mytabley.Fields(CAMPO1)
          found = formateaa(buf, Val(campo3), 0, 0)
          '----------------------------
       End If
       mytabley.Close
       '
       Exit Function
   End If
      If campoz = "TIPOD" Then
          mytabley.Open "SELECT * FROM tipo where  tipo='" & "" & mytablex.Fields("tdocdeli") & "'", cn, adOpenDynamic, adLockOptimistic
       If mytabley.RecordCount > 0 Then 'si existe
       'Set mydby = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
                 '----------------------------
          found = extraer_campos(campoy, CAMPO1, CAMPO2, campo3, campo4, ",")
          buf = "" & mytabley.Fields(CAMPO1)
          found = formateaa(buf, Val(campo3), 0, 0)
          '----------------------------
       End If
       mytabley.Close
       '
       Exit Function
   End If

   If campoz = "CUENTAC" Then
       mytabley.Open "SELECT * FROM cuentac where  local='" & "" & mytablex.Fields("local") & "' and codigo='" & "" & mytablex.Fields("codigo") & "' and tipo='" & "" & mytablex.Fields("tipo") & "' and serie='" & "" & mytablex.Fields("serie") & "' and numero='" & "" & mytablex.Fields("numero") & "'", cn, adOpenDynamic, adLockOptimistic
       If mytabley.RecordCount > 0 Then 'si existe
          '----------------------------
          found = extraer_campos(campoy, CAMPO1, CAMPO2, campo3, campo4, ",")
          buf = "" & mytabley.Fields(CAMPO1)
          found = formateaa(buf, Val(campo3), 0, 0)
          '----------------------------
       End If
       mytabley.Close
       '
       Exit Function
   End If
   If campoz = "CUENTAP" Then
   mytabley.Open "SELECT * FROM cuentap where  local='" & "" & mytablex.Fields("local") & "' and codigo='" & "" & mytablex.Fields("codigo") & "' and tipo='" & "" & mytablex.Fields("tipo") & "' and serie='" & "" & mytablex.Fields("serie") & "' and numero='" & "" & mytablex.Fields("numero") & "'", cn, adOpenDynamic, adLockOptimistic
       If mytabley.RecordCount > 0 Then 'si existe
          '----------------------------
          found = extraer_campos(campoy, CAMPO1, CAMPO2, campo3, campo4, ",")
          buf = "" & mytabley.Fields(CAMPO1)
          found = formateaa(buf, Val(campo3), 0, 0)
          '----------------------------
       End If
       mytabley.Close
       '
       Exit Function
   End If
End If  'fin de busqueda por campos
   
If InStr(campo, ",") > 0 Then   'si es comna
   found = extraer_campos(buf, CAMPO1, CAMPO2, campo3, campo4, ",")
   '------------ver si son seriales
   If Mid$(CAMPO1, 1, 1) = "#" Then      'CREDITO PERSONALES
      CAMPO1 = Mid$(CAMPO1, 2, Len(CAMPO1) - 1)
      buf = "" & mytablex.Fields(CAMPO1)
      If Val(buf) > 0 Then
         found = formateaa("", 1, 2, 0)
         found = formateaa("", 1, 2, 0)
         buf = "      -------------------------"
         found = formateaa(buf, 36, 2, 0)
         buf = "               FIRMA"
         found = formateaa(buf, 36, 0, 0)
      End If
      Exit Function
   End If
   If Mid$(campo, 1, 1) = "!" Then   'COMENTARIOS
   
   End If
   'si es alias de productos
   If Mid$(campo, 1, 1) = ";" Then   'alias
      
   End If
   '--------------------------------------------------------------
   If Mid$(campo, 1, 1) = "?" Then   'si es una condicion exonerado
      CAMPO1 = Mid$(CAMPO1, 2, Len(CAMPO1) - 1)
      buf = "*"
      If Val("" & mytablex.Fields(CAMPO1)) > 0 Then
         buf = ""
      End If
      buf = Mid$(buf, Val(CAMPO2), Val(campo3))
      found = formateaa(buf, Val(campo3), 0, 0)
      Exit Function
   End If
   If Mid$(campo, 1, 1) = "*" Then   'debe restar subtotal-exonerado=subtotal
      CAMPO1 = Mid$(CAMPO1, 2, Len(CAMPO1) - 1)
      sdx = Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("tax"))
      buf = Format(sdx, "0.00")
      buf = Mid$(buf, Val(CAMPO2), Val(campo3))
      found = formateaa(buf, Val(campo3), 0, 0)
      Exit Function
   End If
   If Mid$(campo, 1, 1) = "." Then   'SERIE TICKETERA
      buf = busca_cajay("" & mytablex.Fields("caja"))
      buf = Mid$(buf, Val(CAMPO2), Val(campo3))
      found = formateaa(buf, Val(campo3), 0, 0)
      Exit Function
   End If
   If Mid$(campo, 1, 1) = "@" Then   'Esto numeros a letras
      CAMPO1 = Mid$(CAMPO1, 2, Len(CAMPO1) - 1)
      buf = ""
      buf = pone_letras("" & mytablex.Fields(CAMPO1), "" & mytablex.Fields("moneda"), campo4)
      buf = Mid$(buf, Val(CAMPO2), Val(campo3))
      found = formateaa(buf, Len(buf), 0, 0)
      Exit Function
   End If
   If Mid$(campo, 1, 1) = "$" Then
      CAMPO1 = Mid$(CAMPO1, 2, Len(CAMPO1) - 1)
      buf = "" & mytablex.Fields(CAMPO1)
      If Val(buf) > 0 Then
         If Len("" & mytablex.Fields("tarjeta")) > 0 Then
            buf = Format(Val(buf), "0.00")
            'buf = "Cliente Nro :" & mytablex.Fields("tarjeta") & " Gano " & "" & buf & " Puntos"
            buf = "Gano: " & "" & buf & " Puntos "
            found = formateaa(buf, Val(campo3), 2, 0)
            sdx = busca_cr("1", "" & mytablex.Fields("tarjeta"))
            buf = "Tiene Acumulado " & Format(sdx, "0.00") & " Puntos"
            found = formateaa(buf, 36, 0, 0)
            Exit Function
         End If
         buf = Format(Val(buf), "0.00")
         buf = "Hubiera ganado " & "" & buf & " Puntos"
         found = formateaa(buf, Val(campo3), 0, 0)
      End If
      Exit Function
   End If
   'aqui debe ser seriales
   '--------------------------------------------------------------------------
   Else    'si es :
   If Mid$(campo, 1, 1) = ":" Then
      found = formateaa(campo, Len(campo), 0, 0)
      Exit Function
   End If
   If UCase$(campo) = "TSERVICIO" Then
      If "" & mytablex.Fields("servicio") = "C" Then
         found = formateaa("      " & "Salon." & mytablex.Fields("SALON") & " Mesa " & mytablex.Fields("mesa"), 36, 0, 0)
      End If
      If "" & mytablex.Fields("servicio") = "D" Then
         found = formateaa(" * DOMICILIOS * " & "" & mytablex.Fields("codigo"), 36, 0, 0)
      End If
      If "" & mytablex.Fields("servicio") = "*" Then
         found = formateaa(" *** VENTA RAPIDA *** ", 36, 0, 0)
      End If
      Exit Function
   End If
   If UCase$(campo) = "PONEITEM" Then
      amigohs = Format(flag_contando, "00")
      found = formateaa(amigohs, 2, 0, 0)
      'MsgBox "" & flag_contando
      Exit Function
   End If
   If UCase$(campo) = "DELIVERY" Then
      found = imprime_delivery("" & mytablex.Fields("codigo"))
      Exit Function
   End If
   
   If UCase$(campo) = "PONEMONEDA" Then
      ponemoneda = signopeso
      If "" & mytablex.Fields("moneda") = "S" Then
         ponemoneda = "S/."
      End If
      If "" & mytablex.Fields("moneda") = "D" Then
         ponemoneda = "US$."
      End If
      found = formateaa(ponemoneda, 4, 0, 0)
      Exit Function
   End If
   If UCase$(campo) = "PONEMONEDA1" Then
      ponemoneda = signopeso
      If "" & mytablex.Fields("moneda") = "S" Then
         ponemoneda = "SOLES  "
      End If
      If "" & mytablex.Fields("moneda") = "D" Then
         ponemoneda = "DOLARES"
      End If
      found = formateaa(ponemoneda, 7, 0, 0)
      Exit Function
   End If
   If Mid$(campo, 1, 1) = "." Then   'SERIE TICKETERA
      buf = busca_cajay("" & mytablex.Fields("caja"))
      buf = Mid$(buf, Val(CAMPO2), Val(campo3))
      found = formateaa(buf, Val(campo3), 0, 0)
      Exit Function
   End If
   If Mid$(campo, 1, 1) = "@" Then
      buf = ""
      CAMPO1 = Mid$(CAMPO1, 2, Len(CAMPO1) - 1)
      buf = pone_letras("" & mytablex.Fields(CAMPO1), "" & mytablex.Fields("moneda"), 0)
      buf = Mid$(buf, Val(CAMPO2), Val(campo3))
      found = formateaa(buf, Len(buf), 0, 0)
      Exit Function
   End If
   If Mid$(campo, 1, 1) = "$" Then
      CAMPO1 = Mid$(CAMPO1, 2, Len(CAMPO1) - 1)
      buf = "" & mytablex.Fields(CAMPO1)
      If Val(buf) > 0 Then
         If Len("" & mytablex.Fields("tarjeta")) > 0 Then
            buf = Format(Val(buf), "0.00")
            'buf = "Cliente Nro:" & mytablex.Fields("tarjeta") & " Gano " & "" & buf & " Puntos"
            buf = "Gano: " & "" & buf & " Puntos "
            found = formateaa(buf, Val(campo3), 2, 0)
            sdx = busca_cr("1", "" & mytablex.Fields("tarjeta"))
            buf = "Tiene Acumulado " & Format(sdx, "0.00") & " Puntos"
            found = formateaa(buf, 36, 0, 0)
            'graba en acumulado
            Exit Function
         End If
      buf = Format(Val(buf), "0.00")
      buf = "Hubiera ganado " & "" & buf & " Puntos"
      found = formateaa(buf, Val(campo3), 0, 0)
      End If
      Exit Function
   End If
   buf = "" & mytablex.Fields(campo)
   CAMPO1 = campo
   campo3 = "" & mydbxglo.TableDefs(tablas).Fields(campo).Size
End If
If UCase$(CAMPO1) = "ESTADO" Then
   Dim sdxd As Integer
   buf = "" & mytablex.Fields(CAMPO1)
   If buf = "1" Then
      buf = "*** ANULADO ***"
   Else
      buf = "*** V *** "
   End If
   sdxd = 20
   If Val(CAMPO2) > 0 And Val(campo3) > 0 Then
      buf = Mid$(buf, Val(CAMPO2), Val(campo3))
      sdxd = Val(campo3)
   End If
   found = formateaa(buf, sdxd, 0, 0)
   
   If ascopia = "1" Then
      buf = "   *** COPIA *** "
      found = formateaa(buf, 20, 0, 0)
   End If
   Exit Function
End If
   If Trim(UCase$(CAMPO1)) = "SERIALES" Then
      found = busca_serialesss("" & mytablex.Fields("tipo"), "" & mytablex.Fields("numero"), "" & mytablex.Fields("producto"), contando, CInt(CAMPO2), CInt(campo3))
      If found = 1 Then
         busca_campo1 = "5"
      End If
      Exit Function
   End If
   If Trim(UCase$(CAMPO1)) = "TALLASX" Then
      found = busca_tallasx(mytablex, contando, CInt(CAMPO2), CInt(campo3))
      If found = 1 Then
         busca_campo1 = "5"
      End If
      Exit Function
   End If
   If UCase$(CAMPO1) = "RECETA" Then
    If Len("" & mytablex.Fields("observa")) > 0 Then
       buf = "*" & mytablex.Fields("observa")
       found = verifica_receta_flag1("" & mytablex.Fields("observa"), contando, CInt(CAMPO2), CInt(campo3))
       If found = 0 Then
          'found = formateaa(buf, 35, 2, 0)
       End If
    End If
    Exit Function
   End If
   
If UCase$(campo) = "DELIVERY" Then
      found = imprime_delivery("" & mytablex.Fields("codigo"))
      Exit Function
   End If
'----------------------- -----------------
If Val(CAMPO2) > 0 And Val(campo3) > 0 Then
   buf = Mid$("" & mytablex.Fields(CAMPO1), Val(CAMPO2), Val(campo3))
Else
   buf = "" & mytablex.Fields(CAMPO1)
End If
'If UCase$(CAMPO1) = "REFERENCIA" Then
'   buf = busca_cajay("" & "" & mytablex.Fields("caja"))
'
'End If
'si es campo5 otra forma --------- tipo de letra imprimir---
'-------------------------------------------------------------

Select Case Val("" & mydbxglo.TableDefs(tablabasedatos).Fields(CAMPO1).Type)
      Case 3, 4   'integer
            If campo4 = 1 Then
               buf = Format(Int(Val(buf)), "0")
            End If
            found = formateaa(buf, Val(campo3), 0, 1)
       Case 7  'double
            If campo4 = 0 Then
               buf = Format(Val(buf), "0.00")
               found = formateaa(buf, Val(campo3), 0, 1)
            End If
            If campo4 = 1 Then     'sin decimal pegado a la derecha
               buf = Format(Int(Val(buf)), "0")
               found = formateaa(buf, Val(campo3), 0, 1)
            End If
            If campo4 = 2 Then
               buf = Format(Val(buf), "0.00")
               found = formateaa(buf, Val(campo3), 0, 0)
            End If
            If campo4 = 3 Then
               buf = Format(Int(Val(buf)), "0")
               found = formateaa(buf, Val(campo3), 0, 0)
            End If
            If campo4 = 4 Then 'NORMAL n DECIMALES pegado a la derecha
               buf = Format(Val(buf), "0.00")
               found = formateaa(buf, Val(campo3), 0, 1)
            End If
            If campo4 = 5 Then 'NORMAL n DECIMALES pegado a la derecha
               buf = Format(Val(buf), "0.000")
               found = formateaa(buf, Val(campo3), 0, 1)
            End If
       Case 8
            found = formateaa(buf, 10, 0, 0)
       Case 10
            found = formateaa(buf, Val(campo3), 0, 0)
End Select
Exit Function
cmd9876_err:
MsgBox "Error en busca Campo1 " & campo & error$, 24, "Aviso"
Exit Function
End Function
Sub cerrar_archivo()
On Error GoTo cmd561_err
   Close
   Exit Sub
cmd561_err:
   MsgBox "Aviso en cerrar_archivo " + error$, 48, "Aviso"
   Exit Sub
End Sub
Function extraer_campos(campo As String, CAMPO1 As String, CAMPO2 As String, campo3 As String, campo4 As Integer, Flags As String)
Dim i As Integer
Dim j As Integer
Dim temp As String
i = 0
temp = Trim$(campo)
If Len(temp) = 0 Then Exit Function
Do
   j = InStr(temp, Flags)
   If j > 0 Then
      i = i + 1
      Select Case i
             Case 1: CAMPO1 = Mid$(temp, 1, j - 1)
             Case 2: CAMPO2 = Mid$(temp, 1, j - 1)
             Case 3: campo3 = Mid$(temp, 1, j - 1)
             Case 4: campo4 = CInt(Mid$(temp, 1, j - 1))
             'Case 5: campo5 = Mid$(temp, 1, J - 1)
      End Select
      temp = Trim$(Mid$(temp, j + 1))
     Else
     Exit Function
   End If
Loop
   Exit Function
End Function
Function pone_letras(xrtotal As String, xrmoneda As String, dalongi As Integer) As String
Dim sdx As Double
Dim buf As String
Dim buf1 As String
Dim buf2 As String
Dim found As Integer
Dim ik As Integer
On Error GoTo cmd999999_err

   sdx = Val(xrtotal)
   buf = Format(sdx, "0.00")
   buf1 = Mid$(buf, Len(buf) - 1, 2)
   buf = Mid$(buf, 1, Len(buf) - 3)
   buf = letras(buf, 40)
   buf = LTrim$(Trim$(buf))
   buf = UCase(buf)
   buf2 = LTrim(RTrim(buf)) & " Y " & LTrim(RTrim(buf1))
   If xrmoneda = "D" Then
       buf2 = buf2 & "/100 DOLARES AMERICANOS"
   End If
   If xrmoneda = "S" Then
       buf2 = buf2 & "/100 NUEVOS SOLES"
   End If
   If Len(buf2) < dalongi And dalongi > 0 Then
      For ik = 1 To (dalongi - Len(buf2))
          buf2 = buf2 & " "
      Next ik
   End If
   pone_letras = buf2
   Exit Function
cmd999999_err:
   MsgBox "Error en pone Letras " & error$, 24, "Aviso"
   Exit Function

End Function
Function busca_cr(buf As String, buf1 As String) As Double
End Function
Function busca_serialesss(aatipo As String, aanumero As String, aaproducto As String, contando As Integer, CAMPO2 As Integer, campo3 As Integer)
'
End Function
Function busca_tallasx(mytablex, contando As Integer, CAMPO2 As Integer, campo3 As Integer)
Dim buf As String
Dim mytabley As New ADODB.Recordset
Dim xtallas(17) As String
Dim ytallas(17) As Double
Dim i As Integer
Dim found As Integer
On Error GoTo cmd451213_err


mytabley.Open "select * from linea where linea='" & "" & mytablex.Fields("linea") & "'", cn, adOpenStatic, adLockOptimistic
If mytabley.RecordCount > 0 Then
   For i = 1 To 16
       xtallas(i) = "" & mytabley.Fields("t" & i)
   Next i
End If
mytabley.Close

For i = 1 To 16
  ytallas(i) = Val("" & mytablex.Fields("t" & i))
Next i
buf = ""
For i = 1 To 16
  If ytallas(i) > 0 Then
     buf = buf & xtallas(i) & "/" & ytallas(i) & " "
  End If
Next i
  found = formateaa("", CAMPO2, 0, 0)
  found = formateaa(buf, campo3, 0, 0)
  contando = contando + 1
  busca_tallasx = 1
Exit Function
cmd451213_err:
MsgBox "Busca Seriales:" & error$, 24, "Aviso"
mytablex.Close

Exit Function
End Function

Function verifica_receta_flag(buf As String, contando As Integer)
Dim temp As String
Dim j As Integer
Dim Flags As String
Dim buf1 As String
Dim sw As Integer
Flags = "/"
temp = Trim$(buf)
Dim found As Integer
If Len(temp) = 0 Then Exit Function
temp = Trim$(temp) & "/"
sw = 0
Do
   j = InStr(temp, Flags)
   If j > 0 Then
      buf1 = Mid$(temp, 1, j - 1)
      If Len(buf1) > 0 Then
         buf1 = busca_productoll("" & buf1)
         If Len(buf1) > 0 Then
            If sw = 0 Then
               'found = formateaa("", 1, 2, 0)
               'contando = contando + 1
            End If
            sw = 1
            found = formateaa(buf1, 35, 2, 0)
            contando = contando + 1
            'MsgBox buf1
         End If
         verifica_receta_flag = 1
      End If
      temp = Trim$(Mid$(temp, j + 1))
     Else
     Exit Function
   End If
Loop

End Function
Function verifica_receta_flag1(buf As String, contando As Integer, CAMPO2 As Integer, campo3 As Integer)
Dim temp As String
Dim j As Integer
Dim Flags As String
Dim buf1 As String
Flags = "/"
temp = Trim$(buf)
Dim found As Integer
If Len(temp) = 0 Then Exit Function
temp = Trim$(temp) & "/"
Do
   j = InStr(temp, Flags)
   If j > 0 Then
      buf1 = Mid$(temp, 1, j - 1)
      If Len(buf1) > 0 Then
         buf1 = busca_productoll("" & buf1)
         If Len(buf1) > 0 Then
            found = formateaa("", CAMPO2, 0, 0)
            found = formateaa(buf1, campo3, 2, 0)
            contando = contando + 1
            'MsgBox buf1
         End If
         verifica_receta_flag1 = 1
      End If
      temp = Trim$(Mid$(temp, j + 1))
     Else
     Exit Function
   End If
Loop

End Function
Function busca_cajay(buf As String) As String
Dim mytablex As New ADODB.Recordset


mytablex.Open "select * from parameca where caja='" & "" & buf & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   busca_cajay = "" & mytablex.Fields("serieti")
End If
mytablex.Close


End Function
Function letras(ByVal strnum As String, vlo As Integer) As String
Dim inga As Long
Dim negativo As Variant
Dim L As Integer
Dim una As Variant
Dim millon As Variant
Dim millones As Variant
Dim vez As Integer
Dim maxvez As Integer
Dim k As Integer
Dim strq As String
Dim strb As String
Dim stru As String
Dim strd As String
Dim strc As String
Dim ia As Integer
Dim strn() As String
Dim lo As Integer

  lo = vlo

ReDim unidad(0 To 9) As String
ReDim decena(0 To 9) As String
ReDim centena(0 To 9) As String
ReDim deci(0 To 9) As String
ReDim otros(0 To 15) As String
unidad(1) = "Un"
unidad(2) = "Dos"
unidad(3) = "Tres"
unidad(4) = "Cuatro"
unidad(5) = "Cinco"
unidad(6) = "Seis"
unidad(7) = "Siete"
unidad(8) = "Ocho"
unidad(9) = "Nueve"

decena(1) = "diez"
decena(2) = "veinte"
decena(3) = "treinta"
decena(4) = "cuarenta"
decena(5) = "cincuenta"
decena(6) = "sesenta"
decena(7) = "setenta"
decena(8) = "ochenta"
decena(9) = "noventa"

centena(1) = "ciento"
centena(2) = "doscientos"
centena(3) = "trescientos"
centena(4) = "cuatrocientos"
centena(5) = "quinientos"
centena(6) = "seiscientos"
centena(7) = "setecientos"
centena(8) = "ochocientos"
centena(9) = "novecientos"

deci(1) = "dieci"
deci(2) = "veinti"
deci(3) = "treinta y "
deci(4) = "cuarenta y "
deci(5) = "cincuenta y "
deci(6) = "sesenta y "
deci(7) = "setenta y "
deci(8) = "ochenta y "
deci(9) = "noventa y "

otros(1) = "1"
otros(2) = "2"
otros(3) = "3"
otros(4) = "4"
otros(5) = "5"
otros(6) = "6"
otros(7) = "7"
otros(8) = "8"
otros(9) = "9"
otros(10) = "10"
otros(11) = "once"
otros(12) = "doce"
otros(13) = "trece"
otros(14) = "catorce"
otros(15) = "quince"

On Error GoTo 0
inga = Abs(Val(strnum))
negativo = (inga <> Val(strnum))
strnum = LTrim$(RTrim$(Str$(inga)))
L = Len(strnum)
If inga = 0 Then
   strnum = Left$("cero" & Space$(lo), lo)
   Exit Function
End If
una = True
millon = False
millones = False
If L < 4 Then una = False
If inga > 999999 Then millon = True
If inga > 1999999 Then millones = True
strb = ""
strq = strnum
vez = 0
ReDim strn(1 To 4)
strq = Right$(String$(12, "0") & strnum, 12)

For k = Len(strq) To 1 Step -3
    vez = vez + 1
    strn(vez) = Mid$(strq, k - 2, 3)
Next
maxvez = 4
For k = 4 To 1 Step -1
    If strn(k) = "000" Then
       maxvez = maxvez - 1
       Else
       Exit For
    End If
Next
   
For vez = 1 To maxvez
    stru = "": strd = "": strc = ""
    strnum = strn(vez)
    L = Len(strnum)
    k = Val(Right$(strnum, 2))
    If Right$(strnum, 1) = "0" Then
       k = k \ 10
       strd = decena(k)
    ElseIf k > 10 And k < 16 Then
    k = Val(Mid$(strnum, L - 1, 2))
    strd = otros(k)
    Else
    stru = unidad(Val(Right$(strnum, 1)))
    If L - 1 > 0 Then
       k = Val(Mid$(strnum, L - 1, 1))
       strd = deci(k)
    End If
    End If
    If L - 2 > 0 Then
    k = Val(Mid$(strnum, L - 2, 1))
    strc = centena(k) & " "
    End If
    If stru = "uno" And Left$(strb, 4) = " mil" Then stru = ""
    strb = strc & strd & stru & " " & strb
    If (vez = 1 Or vez = 3) And strn(vez + 1) <> "000" Then strb = " mil " & strb
    If vez = 2 And millon Then
       If millones Then
          strb = " millones " & strb
          Else
          strb = "un millon " & strb
       End If
       End If
    Next
    strb = LTrim$(RTrim$(strb))
    If Right$(strb, 3) = "uno" Then strb = Left$(strb, Len(strb) - 1) & "a"
    'Do
    '  ia = InStr(strb, " ")
    '  If ia = 0 Then Exit Do
    '  strb = Left$(strb, ia - 1) & Mid$(strb, ia + 1)
    'Loop
    If Left$(strb, 6) = "una un " Then strb = Mid$(strb, 5)
    If Left$(strb, 7) = "una mil " Then strb = Mid$(strb, 5)
    If Right$(strb, 16) <> "millones mil una" Then
       ia = InStr(strb, "millones mil una")
       If ia Then strb = Left$(strb, ia + 8) & Mid$(strb, ia + 13)
    End If
    If Right$(strb, 6) = "ciento" Then strb = Left$(strb, Len(strb) - 2)
    If negativo Then strb = "menos " & strb
    strc = Space$(lo)
    LSet strc = strb
    letras = strc
    
End Function
Function busca_productoll(buf As String) As String
Dim mytablex As New ADODB.Recordset
mytablex.Open "select * from producto where producto='" & "" & buf & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   busca_productoll = "" & mytablex.Fields("descripcio")
End If
mytablex.Close


End Function
Sub factura_formato(bxlocal As String, bxtipo As String, bxserie As String, bxnumero As String, ascopia As String)
Dim mytablex As New ADODB.Recordset
Dim mytabley As New ADODB.Recordset
'Dim mytablez As Table
Dim vacu As String

Dim found As Integer
Dim nro_lineas As Integer
Dim contando As Integer
Dim faltante As Integer
Dim i As Integer
Dim archivo_formato As String
On Error GoTo cmd450009_err
       vacu = ""
       'MsgBox "Hola"
       mytablex.Open "select * from tipo where tipo='" & bxtipo & "'", cn, adOpenStatic, adLockOptimistic
       If mytablex.RecordCount = 0 Then
          mytablex.Close
          Exit Sub
       End If
       
       nro_lineas = Val("" & mytablex.Fields("nrolineas"))
       contando = 0
       mytablex.Close
       
       FileName = globaldir & "\temporal\" & gusuario & ".txt"
       found = borra_nombre("" & FileName)
       archivo_formato = buscaformato(bxtipo)
       If Len(archivo_formato) = 0 Then
          MsgBox "factura_formato:No existe archivo Formato ", 48, "Aviso"
          Exit Sub
       End If
       
       'cabeza
       
       mytabley.Open "select * from " & cgusuario & " where local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenStatic, adLockOptimistic
       If mytabley.RecordCount = 0 Then
          mytabley.Close
          Exit Sub
       End If
       'MsgBox ""
       found = proceso_formatos(archivo_formato, mytabley, "{", "}", cgusuario, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
       vacu = "" & mytabley.Fields("acu")
       'MsgBox ""
       '
       'detalle
       flag_contando = 0
       mytablex.Open "select * from " & dgusuariog & " where local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenStatic, adLockOptimistic
       If mytablex.RecordCount > 0 Then
          Do
          If mytablex.EOF Then Exit Do
             flag_contando = contando + 1
             found = proceso_formatos(archivo_formato, mytablex, "/", "\", dgusuariog, "TDETALLE", bxlocal, bxtipo, bxserie, bxnumero, ascopia, contando)
             contando = contando + 1
          mytablex.MoveNext
          Loop
        End If
        mytablex.Close
        
        If bxtipo <> "1" And bxtipo <> "2" And bxtipo <> "5" And bxtipo <> "9" Then
           'If vacu = "V" Or vacu = "F" Or vacu = "P" Or vacu = "D" Or vacu = "I" Or vacu = "S" Then
           'If vacu = "V" Or vacu = "F" Or vacu = "P" Or vacu = "D" Or vacu = "I" Then
           If contando < nro_lineas Then
              For i = contando To nro_lineas
                  Open FileName For Append As #1
                  found = formateaa("", 1, 2, 0)
                  Close #1
              Next i
           'End If
           End If
        End If
        'total  xxxx
       
       'Set mytablex = mydbxglo.OpenTable(cgusuario)
       'mytablex.Index = "tfactura"
       
       found = proceso_formatos(archivo_formato, mytabley, "$", "?", cgusuario, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
       mytabley.Close
       
       '------------------
       Exit Sub
cmd450009_err:
       MsgBox "Aviso Proceso en Compilando " & error$, 24, "Aviso"
       Exit Sub

End Sub
Function borra_nombre(buf As String)
On Error GoTo cmd457_err
   Kill buf
   borra_nombre = 1
   Exit Function
cmd457_err:
   Exit Function
End Function
Function buscaformato(btipo As String) As String
Dim mytablex As New ADODB.Recordset

On Error GoTo cmd4786_err

mytablex.Open "select * from tipo where tipo='" & btipo & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   buscaformato = "" & mytablex.Fields("archivo")
End If
mytablex.Close

Exit Function
cmd4786_err:
MsgBox "Busca Formato, Aviso de Error" & error$, 24, "Aviso"
Exit Function
End Function
Function imprime_archivoj(xbuf As String, xsw, xtipoletra As String)
Dim i As Long
Dim max As Long
Dim buf As String
Dim Free_File As Integer
Dim vr
   On Error GoTo cmd9876_err
   Free_File = FreeFile
   Open xbuf For Input As Free_File
   max = LOF(Free_File)   'numero de letras
   If max < 1 Then
      Close Free_File
      Exit Function
   End If
   Printer.FontName = "courier new"
   'Printer.FontName = "Terminal"
   'Printer.FontName = "Arial"
   If Val(xtipoletra) < 7 Then
      xtipoletra = "9"
   End If
   
   Printer.FontBold = True
   Printer.FontSize = Val(xtipoletra)
   For i = 1 To max
        vr = DoEvents()
        If opcion3 = 1 Then
           If MsgBox("DESEA CANCELAR LA IMPRESION", 1, "AVISO") = 1 Then
              Close Free_File
              Printer.EndDoc
              Exit Function
           End If
           opcion3 = 0
        End If
        Seek Free_File, i
        buf = Input$(1, Free_File)
        If buf = Chr(12) Then
           If xsw = 0 Then
              Printer.NewPage
           End If
           GoTo a2
        End If
        If buf <> Chr(12) Then
           If Chr(10) <> buf Then
                If xsw = 0 Then
                   Printer.Print buf;
                End If
           End If
        End If
a2:
  Next i
  'MsgBox "x"
  Close Free_File
  Printer.EndDoc
  Exit Function
cmd9876_err:
  MsgBox "Error en imprime archivo j " & error$, 48, "Aviso"
  Printer.EndDoc
  Exit Function
End Function
Function Imprime_archivojjx(xbuf As String, xsw As Integer, xtipoletra As String)
Dim i As Long
Dim max As Long
Dim buf As String
Dim Free_File As Integer
Dim vr
Dim k As Integer
   On Error GoTo cmd98716_err
   Free_File = FreeFile
   'MsgBox xbuf
   Open xbuf For Input As Free_File
   max = LOF(Free_File)   'numero de letras
   If max < 1 Then
      Close Free_File
      Exit Function
   End If
   Printer.FontName = "courier new"
   'Printer.FontName = "Terminal"
   'Printer.FontName = "Arial"
   Printer.Print
   If Val(xtipoletra) < 7 Then
      xtipoletra = "9"
   End If
   Printer.FontBold = True
   Printer.FontSize = Val(xtipoletra)
   For i = 1 To max
        'vr = DoEvents()
        'If opcion3 = 1 Then
        '   If MsgBox("DESEA CANCELAR LA IMPRESION", 1, "AVISO") = 1 Then
        '      Close Free_File
        '      Printer.EndDoc
        '      Exit Function
        '   End If
        '   opcion3 = 0
        'End If
        Seek Free_File, i
        buf = Input$(1, Free_File)
        If buf = Chr(12) Then
           If xsw = 0 Then
              Printer.NewPage
           End If
           GoTo a2
        End If
        If buf <> Chr(12) Then
           If Chr(10) <> buf Then
                If xsw = 0 Then
                   Printer.Print buf;
                End If
           End If
        End If
a2:
  For k = 1 To 2000
  Next k
  Next i
  'MsgBox "x"
  Close Free_File
  Printer.EndDoc
  Exit Function
cmd98716_err:
  MsgBox "Error en imprime archivo jj " & error$, 48, "Aviso"
  Printer.EndDoc
  Exit Function

End Function
Function imprime_archivojj(path As String, xsw As Integer, xtipoletra As String)
       
    Dim Free_File As Integer
    Dim datos As String
    Dim pos As Integer
    Dim L As String
    Dim i As Integer
    Dim Palabra As String
    Dim vbcrlf As String
    Dim buf As String
    On Error GoTo cmd89000_err
      
    ' número de archivo libre
    Free_File = FreeFile
    vbcrlf = Chr$(10) + Chr$(13)
       
    ' abre el archivo para leerlo
    'MsgBox path
    Open path For Input As Free_File
      
    ' Almacena los datos del archivo en la variable
    datos = Input(LOF(Free_File), Free_File)
    ' cierra el archivo
         'MsgBox printer.FontName
         'printer.FontName = "15 cpi"  'solo es para
         'printer.FontName = "Currer New"
         'printer.FontSize = 9
         'printer.FontBold = False
         
         '-----se adiciono del anhetior..
         Printer.FontName = "courier new"
         'Printer.FontName = "Terminal"
         'Printer.FontName = "Arial"
         Printer.Print
         If Val(xtipoletra) < 7 Then
         xtipoletra = "9"
         End If
         Printer.FontBold = True
         Printer.FontSize = Val(xtipoletra)
         '------------------------------
         
         'If Len(tipoleta) > 0 Then
         '   Printer.FontName = tipoleta  'solo es para
         'End If
         'If Val(tamano) > 0 Then
         '   Printer.FontSize = Val(tamano)
         '   'MsgBox ""
         'End If
         'If negrita = "S" Then
         '   Printer.FontBold = True
         'End If

    Close Free_File
    Do While Len(datos) > 0
           
        pos = InStr(datos, vbcrlf)
        If pos = 0 Then
            L = datos
            datos = ""
        Else
            ' linea
            L = Left$(datos, pos - 1)
               
            datos = Mid$(datos, pos + 2)
        End If
       
    ' palabras
    Do While Len(L) > 0
        ' posición para extraer la palabra
        pos = InStr(L, " ")
        If pos = 0 Then
            Palabra = L
            L = ""
        Else
            Palabra = Left$(L, pos)
            L = Mid$(L, pos + 1)
        End If
       
    Printer.Print Palabra;
    Loop
    'printer.Print Chr$(27) + "i";   'epson
    'For i = 1 To 10
    '   printer.Print
    'Next i
    Printer.Print
    Loop
    'MsgBox ""
    'buf = Chr$(27) + "i"  'epson
    'printer.Print buf;
    
    'buf = Chr(27) & Chr(105)
    'printer.Print buf;
    
    'printer.Print
    Printer.EndDoc
    Exit Function
cmd89000_err:
    MsgBox "Aviso en imprimir_archivo ,no es el Driver correcto Impresora " + error$, 48, "Aviso"
    Exit Function

  
End Function

Function busca_combox(xxyzcontrol As Control, buf As String)
On Error GoTo cmd45_err
Dim i As Integer
Dim sw As Integer
sw = 0
For i = 0 To xxyzcontrol.ListCount - 1
   If xxyzcontrol.List(i) = buf Then
      busca_combox = i
      sw = 1
      Exit For
   End If
Next i
If sw = 0 Then
   busca_combox = 0
End If
Exit Function
cmd45_err:
busca_combox = 0
Exit Function
End Function

Function calcula_saldo(sdx As Double, sdx1 As Double) As String
Dim buf As String
Dim dsdx As Double
Dim rsdx As Double
Dim signo1 As Variant
Dim buf1 As String
    '--------------------
    If extraer_decimal(sdx1) > 0 Then  'si es decimal dejar como esta
       calcula_saldo = "" & sdx
       Exit Function
    End If
    If sdx = 0 Then
       calcula_saldo = "0"
       Exit Function
    End If
    If Int(sdx1) = 0 Then
       calcula_saldo = "" & sdx
       Exit Function
    End If
    '--------------------
    sdx = Val(Format(sdx, "0.00"))
    buf1 = ""
    If sdx < 0 Then
        buf1 = "-"
    End If
    If sdx1 = 1 Or sdx <= 1 Then
       calcula_saldo = "" & sdx
       'MsgBox buf1 & Str(sdx)
       Exit Function
    End If
    buf = ""
    sdx = Abs(sdx)
    If sdx1 = 0 Then
       calcula_saldo = "0"
       Exit Function
    End If
    dsdx = sdx Mod sdx1
    rsdx = (sdx - dsdx) / sdx1
    dsdx = sdx - rsdx * sdx1
    buf = Format(dsdx, "0")
    If Val(buf) <= 99 And Val(buf) >= 10 Then
       buf = "0" & buf
    End If
    If Val(buf) <= 9 Then
       buf = "00" & buf
    End If
    If Val(buf) = 0 Then
       buf = buf1 & "" & rsdx
       Else
       buf = buf1 & "" & rsdx & "." & buf
    End If
    buf = Format(Val(buf), "0.000")
    calcula_saldo = buf
End Function
Function extraer_decimal(sdx As Double)
Dim sdx1 As Double
Dim buf As String
buf = Format(sdx, "#.#")
If InStr(buf, ".") Then
extraer_decimal = Val(Mid$(buf, InStr(buf, ".") + 1, 2))
Else
extraer_decimal = 0
End If
End Function
Function etiqueta_disco() As String
Dim unidad As String
Dim cad1 As String * 256
  Dim cad2 As String * 256
  Dim numSerie As Long
  Dim longitud As Long
  Dim flag As Long
  unidad = "c:\"
  Call GetVolumeInformation(unidad, cad1, 256, numSerie, longitud, flag, cad2, 256)
  etiqueta_disco = cad1
End Function
Function existe_tarjeta_sonido()
Dim inf As Integer
inf = waveOutGetNumDevs()
If inf > 0 Then
MsgBox "Tarjeta de sonido soportada.", vbInformation, "Informacion: Tarjeta de sonido"
Else
MsgBox "Tarjeta de sonido no soportada.", vbInformation, "Informacion: Tarjeta de sonido"
End If
End

End Function
Function permite_entrada_salida(buf As String)
Select Case buf
       'valida si permite movimientos de almacen
       Case "A", "B", "C", "D", "G", "E", "F", "J", "K", "L", "M", "P", "N", "0", "S", "T"
       permite_entrada_salida = 1
End Select
End Function
Function extra_loquesea(buf As String) As String
Dim j
Dim buf1 As String
buf1 = ""
If InStr(buf, "|") > 0 Then
   j = InStr(buf, "|")
   buf1 = Mid$(buf, 1, j - 1)
   Else
   buf1 = buf
End If
extra_loquesea = buf1
End Function
Sub cabecera_tipico(buf1 As String, buf2 As String, buf3 As String)
Dim buf As String
Dim found As Integer
Dim mytablex As New ADODB.Recordset

   
   buf = ""
   
   mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount > 0 Then
      buf = buf & "" & mytablex.Fields("cabecera1")
   End If
   mytablex.Close
   
   found = formateaa(buf, 70, 0, 0)
   buf = formateaa(" ", 5, 0, 0)
   buf = "Pagina :" & contpag
   found = formateaa(buf, 15, 2, 0)
   
   buf = "Usuario :" & buf3
   found = formateaa(buf, 70, 0, 0)
   buf = formateaa(" ", 5, 0, 0)
   buf = "Hora   :" & Format(Now, "hh:mm:ss")
   found = formateaa(buf, 20, 2, 0)
   buf = formateaa("", 1, 2, 0)
End Sub
Function crypt(pw, cryptee)
Dim i
Dim pchar
Dim cchar
    Do While Len(pw) < Len(cryptee)
        pw = pw & pw
    Loop
    For i = 1 To Len(cryptee)
        pchar = Asc(Mid$(pw, i, 1))
        cchar = Asc(Mid$(cryptee, i, 1))
        Mid$(cryptee, i, 1) = Chr$(pchar Xor cchar)
    Next i
    crypt = cryptee
    

End Function
Function valida_fecha(buf As String)
Dim buf1 As String
If Len(buf) <> 10 Then
   Exit Function
End If
If Not IsNumeric(Mid$(buf, 1, 2)) Then
   Exit Function
End If
If Not IsNumeric(Mid$(buf, 4, 2)) Then
   Exit Function
End If
If Not IsNumeric(Mid$(buf, 7, 4)) Then
   Exit Function
End If
If Val(Mid$(buf, 1, 2)) < 1 And Val(Mid$(buf, 1, 2)) > 31 Then
   Exit Function
End If
If Val(Mid$(buf, 4, 2)) < 1 And Val(Mid$(buf, 4, 2)) > 12 Then
   Exit Function
End If

If IsDate(buf) Then
   valida_fecha = 1
End If
End Function

Function valida_hora(buf As String)
If Mid$(buf, 3, 1) = ":" Then
If Val(Mid$(buf, 1, 2)) >= 0 And Val(Mid$(buf, 1, 2)) <= 24 Then
   If Val(Mid$(buf, 4, 2)) >= 0 And Val(Mid$(buf, 4, 2)) <= 60 Then
   valida_hora = 1
   End If
End If
End If
End Function
Function busca_paridadg(sw As Integer) As Double
On Error GoTo cmd7_err
Dim mytablex As New ADODB.Recordset
mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   If sw = 0 Then
      busca_paridadg = Val("" & mytablex.Fields("parivta"))
   End If
   If sw = 1 Then
      busca_paridadg = Val("" & mytablex.Fields("paricomp"))
   End If
End If
'------------------------------------- ------------
mytablex.Close
Exit Function
cmd7_err:
Exit Function

End Function
Sub ValidarFecha(fecha As String, valida As Boolean)

Dim cadena As Date
On Error GoTo error
cadena = Format(fecha, "dd/mm/yyyy")
If Not IsDate(cadena) Then
    MsgBox "Compruebe que ha introducido bien la fecha.", vbInformation
    Exit Sub
End If
If cadena > Date Then
    valida = True
    GoTo error
Else
    valida = False
End If
    Exit Sub
error:
MsgBox "La fecha no puede ser posterior a la fecha de hoy.", vbInformation, "Fecha inválida"
    valida = True
    Exit Sub
End Sub

 Sub PrintTXTFile(FileName As String)
 'imprimir un archivo de texto
          Dim X As Integer
          Dim s As String
          X = FreeFile
          On Error GoTo HandleError
          Open FileName For Input As X
          Do While Not EOF(X)
              Line Input #X, s
              Printer.Print s
          Loop
          Printer.EndDoc
          Close #X
          Exit Sub
HandleError:
          MsgBox "Error :" & Err.Description, vbCritical, "Imprimiendo fichero..."
End Sub
 Sub ExecuteCommand(FileToExecute As String)
On Error GoTo OpenError
Dim Lng As Long
Lng = Shell(FileToExecute, vbNormalFocus)
OpenError:
If Err.Number <> 0 Then
   MsgBox "Cannot Understand Message! ", vbOKOnly, "Help"
   Resume Next
End If
End Sub
Function copia_tmpweb()
On Error GoTo cmd235_err
FileCopy globalweb & "\temp\wfactura.dbf", globalweb & "\r\" & "_c" & gusuario & ".dbf"
FileCopy globalweb & "\temp\wfactura.cdx", globalweb & "\r\" & "_c" & gusuario & ".cdx"
FileCopy globalweb & "\temp\wdetalle.dbf", globalweb & "\r\" & "_d" & gusuario & ".dbf"
FileCopy globalweb & "\temp\wdetalle.cdx", globalweb & "\r\" & "_d" & gusuario & ".cdx"
FileCopy globalweb & "\temp\wfpagov.dbf", globalweb & "\r\" & "_f" & gusuario & ".dbf"
FileCopy globalweb & "\temp\wfpagov.cdx", globalweb & "\r\" & "_f" & gusuario & ".cdx"
copia_tmpweb = 1
Exit Function
cmd235_err:
Exit Function
End Function
Function imprime_linea(slpt As String, b As String)

End Function
Function imprime_puerto_serial(slpt As String)

Exit Function
End Function
Function star_sp342(puerto As String, sw As Integer) 'sw determina tipo de dispositivo
Dim r%
Dim s As Boolean
Dim found As Integer
Dim buf As String
Dim vr As Integer
Dim bufferr As String
Dim slpt As String
Dim i As Long
Dim j As Integer
Dim max As Long
Dim velox As String
Dim lngStatus As Long
Dim X As Integer
Dim strError  As String
Dim nficsal As Integer
Dim b As String
Dim contin As Integer
On Error GoTo cmd13081_err
If Len(FileName) = 0 Then
   MsgBox "Mensaje,Nombre archivo no existe ", 24, "Aviso"
   Exit Function
End If
   i = 0
   cerrar_archivo
   cerrar_puertoscom
   slpt = puerto
   If puerto = "BAR" Then
      impresion_codbar "IMPRIME"
      Exit Function
   End If
   If slpt = "COM1" Or slpt = "COM2" Or slpt = "COM3" Or slpt = "COM4" Or slpt = "COM5" Then
          imprime_puerto_serial slpt
          cerrar_archivo
          Exit Function
   End If
   'MsgBox slpt
   'End
   Open FileName For Input As #1
   Open slpt For Output As #2
   'MsgBox Filename
   'End
   max = LOF(1)   'numero de letras
   'MsgBox max
   'End
   For i = 1 To max
        Seek #1, i
        buf = Input$(1, #1)
        Print #2, buf;
   Next i
   'MsgBox ""
   'End
   '---------------------------------------------
   Close #1
   Close #2
   cerrar_archivo
Exit Function
cmd13081_err:
MsgBox "Error en impresora " & i & " " & error$, 1, "Mensaje"
   Close #1
   Close #2
   Exit Function
End Function
Function corte_papel(apuerto As String, sw As Integer)
Dim buf As String
Dim i
Dim found As Integer


Select Case apuerto
       Case "LPT1", "LPT2", "LPT3", "LPT4", "LPT5"
            If sw = 0 Then   'star
               'buf = Chr$(27) + "d" + Chr$(0) 'star sp342
               i = FreeFile
               Open apuerto For Output As i
               'buf = Chr$(28) + Chr$(29)  'star sp200
               'buf = Chr$(7)   'star
               buf = Chr$(28) + Chr$(29)  'star sp200
               Print #i, buf;
               Close i
            End If
            If sw = 1 Then   'epson
               i = FreeFile
               Open apuerto For Output As i
               'buf = Chr$(27) + "i"  'epson
               'buf = Chr$(27) + "p" + Chr$(0) + Chr$(25) + Chr$(250) 'EPSON
               buf = Chr$(27) + "i"  'epson
               Print #i, buf;
               Close i
            End If
       Case "BAR"
            If sw = 0 Then
               buf = Chr$(28) + Chr$(29)  'star sp200
               impresion_codbar buf
            End If
            If sw = 1 Then
               buf = Chr$(27) + "i"  'epson
               impresion_codbar buf
            End If
            
       Case "COM1", "COM2", "COM3", "COM4", "COM5"
            If sw = 0 Then   'star
               'buf = Chr$(27) + "d" + Chr$(0) 'star sp342
               buf = Chr$(28) + Chr$(29)  'star sp200
               found = imprime_linea(apuerto, buf)
            End If
            If sw = 1 Then   'epson
               buf = Chr$(27) + "i"  'epson
               found = imprime_linea(apuerto, buf)
            End If
End Select


End Function
Function abre_puerto(apuerto As String, sw As Integer)  'solo gaveta dinero
Dim buf As String
Dim found As Integer
Dim i
On Error GoTo cmd8912_err
Select Case apuerto
       Case "LPT1", "LPT2", "LPT3", "LPT4", "LPT5"
            If sw = 0 Then   'star
               'buf = Chr$(27) + "d" + Chr$(0) 'star sp342
               i = FreeFile
               Open apuerto For Output As i
               buf = Chr$(28) + Chr$(29)  'star sp200
               buf = Chr$(7)   'star
               Print #i, buf;
               Close i
            End If
            If sw = 1 Then   'epson
               i = FreeFile
               Open apuerto For Output As i
               buf = Chr$(27) + "i"  'epson
               buf = Chr$(27) + "p" + Chr$(0) + Chr$(25) + Chr$(250) 'EPSON
               Print #i, buf;
               Close i
            End If
       Case "BAR"
            If sw = 0 Then
               buf = Chr$(28) + Chr$(29)  'star sp200
               buf = Chr$(7)   'star
               impresion_codbar buf
            End If
            If sw = 1 Then
               buf = Chr$(27) + "i"  'epson
               buf = Chr$(27) + "p" + Chr$(0) + Chr$(25) + Chr$(250) 'EPSON
               impresion_codbar buf
            End If
            
       Case "COM1", "COM2", "COM3", "COM4", "COM5"
            If sw = 0 Then   'star
               
               'buf = Chr$(27) + "d" + Chr$(0) 'star sp342
               buf = Chr$(28) + Chr$(29)  'star sp200
               buf = Chr$(7)   'star
               found = imprime_linea(apuerto, buf)
               
            End If
            If sw = 1 Then   'epson
               buf = Chr$(27) + "i"  'epson
               buf = Chr$(27) + "p" + Chr$(0) + Chr$(25) + Chr$(250) 'EPSON
               found = imprime_linea(apuerto, buf)
            End If
End Select
Exit Function
cmd8912_err:
MsgBox "Error en Abre Cajon " + error$, 48, "Aviso"
Exit Function
End Function
Function valida_ruc(as_numero2 As String)
Dim ic_num As Integer
Dim as_numero1 As String
Dim xx As Integer
Dim as_numero As String
'valida_ruc = 1
'Exit Function
Dim lc_compara As Integer
Dim lc_residuo As Integer
Dim lc_valida As Integer
If Len(as_numero2) <> 11 Then Exit Function
If Mid$(as_numero2, 1, 2) <> "20" Then
   valida_ruc = 1
   Exit Function
End If
as_numero = as_numero2
If Mid$(as_numero2, 1, 2) = "20" Then
   as_numero = Mid$(as_numero2, 3, 8)
End If
xx = Len(as_numero)
If Len(as_numero) <> 8 Then Exit Function
as_numero1 = as_numero
If Len(as_numero) = 11 Then
   as_numero1 = Mid$(as_numero, 3, 8)
End If
If Not IsNumeric(as_numero1) Then Exit Function
ic_num = Val(Mid$(as_numero1, 1, 1)) * 2 + Val(Mid$(as_numero1, 2, 1)) * 7 + Val(Mid$(as_numero1, 3, 1)) * 6 + Val(Mid$(as_numero1, 4, 1)) * 5 + Val(Mid$(as_numero1, 5, 1)) * 4 + Val(Mid$(as_numero1, 6, 1)) * 3 + Val(Mid$(as_numero1, 7, 1)) * 2
lc_compara = Val(Mid$(as_numero1, 8, 1))
lc_residuo = ic_num Mod 11
lc_valida = Int(11 - lc_residuo)
If lc_valida > 9 Then lc_valida = lc_valida - 10
If lc_valida <> lc_compara Then Exit Function
valida_ruc = 1
End Function
Function copiar_temporalxxx()
On Error GoTo cmd23_err
FileCopy globaldat & "\controld.dbf", globaldat & "\" & "__" & gusuario & ".dbf"
copiar_temporalxxx = 1
Exit Function
cmd23_err:
Exit Function

End Function
Function imprime_delivery(buf1 As String)
On Error GoTo cmd8912_err
Dim buf As String


Dim mytabley As New ADODB.Recordset
Dim found As Integer

found = formateaa("DATOS DELIVERY", 30, 2, 0)

mytabley.Open "select * from clientes where codigo='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic
If mytabley.RecordCount > 0 Then
          buf = "Telef:" & mytabley.Fields("Telefono")
          found = formateaa(buf, 20, 2, 0)
          If Len(Mid$("" & mytabley.Fields("nombre"), 1, 30)) > 0 Then
             buf = Mid$("Nom:" & mytabley.Fields("nombre"), 1, 34)
             found = formateaa(buf, 30, 2, 0)
          End If
          If Len(Mid$("" & mytabley.Fields("nombre"), 31, 60)) > 0 Then
             buf = Mid$("" & mytabley.Fields("nombre"), 31, 30)
             found = formateaa(buf, 30, 2, 0)
          End If
          If Len(Mid$("" & mytabley.Fields("Direccion"), 1, 30)) > 0 Then
             buf = Mid$("Dir:" & mytabley.Fields("Direccion"), 1, 34)
             found = formateaa(buf, 30, 2, 0)
          End If
          If Len(Mid$("" & mytabley.Fields("Direccion"), 31, 60)) > 0 Then
             buf = Mid$("" & mytabley.Fields("Direccion"), 31, 30)
             found = formateaa(buf, 30, 2, 0)
          End If
          buf = "" & mytabley.Fields("Distrito")
          found = formateaa(buf, 30, 2, 0)
          If Len(Mid$("" & mytabley.Fields("observa"), 1, 30)) > 0 Then
             buf = Mid$("Ref:" & mytabley.Fields("observa"), 1, 34)
             found = formateaa(buf, 30, 2, 0)
          End If
          If Len(Mid$("" & mytabley.Fields("observa"), 31, 60)) > 0 Then
             buf = Mid$("" & mytabley.Fields("observa"), 31, 30)
             found = formateaa(buf, 30, 2, 0)
          End If
          '----------------------------
       End If
       mytabley.Close
       
       Exit Function
cmd8912_err:
      MsgBox "Error en imprime_deliveri " + error$, 48, "Aviso"
       mytabley.Close
       
End Function
Function valida_red()
Dim lDrive As Long
Dim szRoot As String
szRoot = Mid$(App.path, 1, 3)
'MsgBox szRoot
'MsgBox szRoot
'Poner aquí la unidad del CD-ROM o la que queramos comprobar
lDrive = GetDriveType(szRoot)
'MsgBox lDrive

If lDrive = 4 Then
   'MsgBox "Hola"
    valida_red = 1
End If
End Function
Function impresion_codbar1(buf As String)
Dim objPrinter
Dim NombreImpresora_sp As String
Dim NombreImpresora_us As String
On Error GoTo cmd78911_err
NombreImpresora_us = "Genérico / sólo texto"
NombreImpresora_sp = "Argox X-1000v PPLA"
    Set objPrinter = New PrinterAPI.clsPrinter
    'Intenta Seleccionar la Impresora
    MsgBox Printer.DeviceName
    
    NombreImpresora_sp = Printer.DeviceName
    
    If (objPrinter.SetPrinter(NombreImpresora_sp) = False) Then
        If (objPrinter.SetPrinter(NombreImpresora_us) = False) Then
            MsgBox "No se Encuentra instalada la Impresora " & NombreImpresora_sp & "o " & NombreImpresora_us, vbInformation
            Exit Function
        End If
    End If
    objPrinter.PrintData buf
    objPrinter.PrintEndDoc
    Set objPrinter = Nothing
    Exit Function
cmd78911_err:
    MsgBox "Reinicie la impresora " + error$, 48, "Aviso"
    Exit Function
End Function
Function impresion_codbar(buf As String)
Dim objPrinter
Dim max As Long
Dim i As Long

Dim NombreImpresora_sp As String
Dim NombreImpresora_us As String
On Error GoTo cmd178911_err
NombreImpresora_us = "Genérico / sólo texto"
NombreImpresora_sp = "Argox X-1000v PPLA"
    Set objPrinter = New PrinterAPI.clsPrinter
    'Intenta Seleccionar la Impresora
    NombreImpresora_sp = Printer.DeviceName
    If (objPrinter.SetPrinter(NombreImpresora_sp) = False) Then
        If (objPrinter.SetPrinter(NombreImpresora_us) = False) Then
            MsgBox "No se Encuentra instalada la Impresora " & NombreImpresora_sp & "o " & NombreImpresora_us, vbInformation
            Exit Function
        End If
    End If
      
      If buf = "IMPRIME" Then
         Open FileName For Input As #1
         buf = ""
         max = LOF(1)
         For i = 1 To max
             Seek #1, i
             buf = Input$(1, #1)
             objPrinter.PrintData buf
             Sleep (500)
         Next i
         Close #1
         cerrar_archivo
         objPrinter.PrintData buf
         objPrinter.PrintEndDoc
         Set objPrinter = Nothing
         Exit Function
      End If
    objPrinter.PrintData buf
    objPrinter.PrintEndDoc
    Set objPrinter = Nothing
    Exit Function
cmd178911_err:
    MsgBox "Reinicie la impresora " + error$, 48, "Aviso"
    Exit Function

End Function
Sub cerrar_puertoscom()

End Sub
Function copiandox(buf As String, buf1 As String)
    On Error GoTo ErrHandler1
    FileCopy buf, buf1
    copiandox = 1
    Exit Function
ErrHandler1:
   MsgBox "Error al inicializa...avise al adm " + error$, 48, "Aviso"
   Exit Function
End Function
Sub copiando(buf As String, buf1 As String)
Dim DestFile, Msg   ' Declare variables.
    On Error GoTo ErrHandler
    FileCopy buf, buf1
    Exit Sub
ErrHandler:
    If Err = 55 Then    ' File already open.
        MsgBox "Tablas ya abiertas ,Salga y Vuelva a Ingresar. O limpie Temporales ", 24, "Aviso"
        End
    Else
        MsgBox "Por Favor Limpiar Temporales  Limpiar Temporales .", 48, "Aviso"
        End
    End If
    Resume Next
End Sub
