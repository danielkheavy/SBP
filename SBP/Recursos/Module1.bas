Attribute VB_Name = "Module1"
Option Explicit
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function GetComputerName _
                Lib "kernel32" _
                Alias "GetComputerNameA" (ByVal lpBuffer As String, _
                                          nSize As Long) As Long
Declare Function FindWindow _
        Lib "user32" _
        Alias "FindWindowA" (ByVal lpClassName As String, _
                             ByVal lpWindowName As String) As Long
Declare Function SetWindowPos _
        Lib "user32" (ByVal hwnd As Long, _
                      ByVal hWndInsertAfter As Long, _
                      ByVal X As Long, _
                      ByVal Y As Long, _
                      ByVal cX As Long, _
                      ByVal cY As Long, _
                      ByVal wFlags As Long) As Long
Declare Function GetActiveWindow Lib "user32" () As Long

Private Declare Function CreateFile _
                Lib "kernel32" _
                Alias "CreateFileA" (ByVal lpFileName As String, _
                                     ByVal dwDesiredAccess As Long, _
                                     ByVal dwShareMode As Long, _
                                     ByVal lpSecurityAttributes As Long, _
                                     ByVal dwCreationDisposition As Long, _
                                     ByVal dwFlagsAndAttributes As Long, _
                                     ByVal hTemplateFile As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Sub RtlMoveMemory _
        Lib "kernel32" (lpvDest As Any, _
                        lpvSource As Any, _
                        ByVal cbCopy As Long)

'Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)
Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long

Public Declare Function GetVolumeInformation& _
               Lib "kernel32" _
               Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
                                              ByVal pVolumeNameBuffer As String, _
                                              ByVal nVolumeNameSize As Long, _
                                              lpVolumeSerialNumber As Long, _
                                              lpMaximumComponentLength As Long, _
                                              lpFileSystemFlags As Long, _
                                              ByVal lpFileSystemNameBuffer As String, _
                                              ByVal nFileSystemNameSize As Long)
Declare Function GetDriveType _
        Lib "kernel32" _
        Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

'997061114
Private Const OF_READ = &H0&

Private lpFSHigh As Long

Private Declare Function lOpen _
                Lib "kernel32" _
                Alias "_lopen" (ByVal lpPathName As String, _
                                ByVal iReadWrite As Long) As Long

Private Declare Function lclose _
                Lib "kernel32" _
                Alias "_lclose" (ByVal hFile As Long) As Long

Private Declare Function GetFileSize _
                Lib "kernel32" (ByVal hFile As Long, _
                                lpFileSizeHigh As Long) As Long

Public Type bdvisor

    ppuerto As String * 30
    vvelocidad As String * 30
    mmensaje1 As String * 30
    mmensaje2 As String * 30

End Type

Global licencia_remoto As String

Public Type ipmaquina

    local1 As String * 2
    base As String * 20
    clave As String * 20
    ip As String * 20
    defecto As String * 1
    nombre As String * 30

End Type

'01/06/2017  KENYO APARECE EN CONGELA NOMBRE DE CLIENTE PARA DESCONGELAR
Global opcioncongela   As Integer

Global nombre_sistema1 As String

'01/06/2017  KENYO APARECE EN CONGELA NOMBRE DE CLIENTE PARA DESCONGELAR

''11/07/2017 kenyo multicomandas
Global nroimpresion    As Integer

''11/07/2017 kenyo multicomandas

Global swptovta        As String

Global globalmesero    As String

Global codigohuella    As String

Global nrodecimal      As String

Global sw_acura        As Integer

Global rrlocal11       As String

Global rrtipo          As String

Global rrserie         As String

Global rrnumero        As String

''' kenyo 31/08/2017 Modulo delivery personalizado
Global tipod           As String

Global documentod      As String

Global seried          As String

Global numerod         As String

''' kenyo 31/08/2017 Modulo delivery personalizado

'16/03/2018 No sale error en Ingreso desde Menu
Global AbreGaveta      As String

'16/03/2018 No sale error en Ingreso desde Menu

Global anticipoo       As String

Global sw_consulta     As Integer

Global nombre_sistema  As String

Global clave_servidor  As String

Global VISTA           As String

Global ver_xproducto   As String

Global ejecutawor      As Integer

Global opciontablet    As String

Global stockvirtual    As String

Global dbclie          As New ADODB.Recordset  'ojo es general

Global flag_denisse    As String

Global glomesa         As String

Global flag_comanda    As String

Global serial_number   As String

Public Const SWP_HIDEWINDOW = &H80

Public Const SWP_SHOWWINDOW = &H40

Global tipodescuento    As String

Global valordescuento   As Double

Global glocal           As String

Global txtotlare        As Double

Global mysnap           As Snapshot

Global tipoletra        As String

Global opcion5          As Integer

Global orionv4          As String

Global suma1            As Double

Global suma2            As Double

Global suma3            As Double

Global suma4            As Double

Global suma5            As Double

Global suma6            As Double

Global suma7            As Double

Global suma8            As Double

Global suma9            As Double

' ''10/10/2017 Reporte de Seguimiento de facturas En Excel
Global suma10           As Double

Global fin              As Integer

' ''10/10/2017 Reporte de Seguimiento de facturas En Excel

Global ssuma1           As Double

Global ssuma2           As Double

Global ssuma3           As Double

Global ssuma4           As Double

Global ssuma5           As Double

Global ssuma6           As Double

Global ssuma7           As Double

Global ssuma8           As Double

Global ssuma9           As Double

Global amsw             As Integer

Global mensaje_print    As String

Global xxxsoles         As String

'Global unidadx As String

Global opcion1          As String

Global opcion2          As String

Global opcion3          As String

Global signopeso        As String

Global mensaje_bloqueo  As String

Global globaldir        As String

Global globaldat        As String

Global gglobaldir       As String

Global gglobaldat       As String

Global globalpath       As String

Global globalweb        As String

Global globalemp        As String

Global globalocal       As String

Global empresapos       As String

Global mytablexxx       As Table

Global mydbxxx          As Database

Global globalpri        As String

Global globalcont       As String

Global gusuario         As String

Global ngusuario        As String

Global FileName         As String

Global ncanal           As Integer

Global mydb11           As Database

'Global mytable11 As Table
Global mydbxglo         As Database

Global mydbzglo         As Database 'contable

Global flag_contando    As Integer   'flag de nro item en el formato factura

Global dbserial         As String

Global xarchivo         As String

Global xarchivo1        As String

Global dia              As String

Global tipo_servicio    As String

Global cgusuario        As String

Global usuariopos       As String

Global contlin          As Double

Global contpag          As Double

Global ticketera_cajon  As Integer

Global sgusuario        As String 'servicio tecnico

Global dgusuario        As String

Global dgusuariog       As String

Global fgusuario        As String

Global fpusuario        As String

Global fpusuarior       As String

Global gocabeza         As String

Global godetalle        As String

Global gofpago          As String

Global dbbase           As String

Global dbca             As String

Global dbing            As String

Global dbde             As String

Global dbfp             As String

Global dbtalla          As String

Global xnpuerto         As String

Global xnpuerto1        As String

Global flag_clave1      As Integer

Public vHabitacion      As String

Global vIMPRIMIR        As Integer

Global vBUSCAXPRODUCTO  As Integer

'inicio 30/05/2017 pll
Global nombre           As String

'fin 30/05/2017 pll

'Color por familia y producto  30/05/2018

Global ffamilia         As String

'Color por familia y producto  30/05/2018

'' 13/01/2018 Fecha de registro y cambio de congelados
Public fechaicongela    As String
'' 13/01/2018 Fecha de registro y cambio de congelados

'''' 17/07/2018 Factura de Exportación
Global my_tipooperacion As String

Global my_tipoigv       As String

'''' 17/07/2018 Factura de Exportación
Sub Main()

    'tsplas.Show 1
End Sub

Function busca_tipoprecio() As String

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM parame where codigo='01'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        busca_tipoprecio = Trim("" & mytablex.Fields("tipoprecio"))

    End If

    mytablex.Close

End Function

Function busca_tiporpt(buf As String, sw As Integer) As String

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM tipo where  tipo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        busca_tiporpt = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close

End Function

Function gaveta_cola(xxpuerto As String)

    'ACA IMPRIME
    Dim oldprinter

    oldprinter = Printer.DeviceName
    selecciona_impresoras (xxpuerto)
    Printer.Print ""
    Printer.EndDoc
    selecciona_impresoras (oldprinter)
    'MsgBox Printer.DeviceName
    gaveta_cola = 1

End Function

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

    Dim sdx  As Integer

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

        If sw = 2 Then
            buf1 = Trim(buf)

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

    On Error GoTo handler

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
handler:

End Function

Function serie_disco_duro() As String

    Dim unidad As String

    Dim buf    As String

    Dim xbuf   As String

    Dim ybuf   As String

    On Error GoTo cmd9011_err

    Dim cad1     As String * 256

    Dim cad2     As String * 256

    Dim numSerie As Long

    Dim longitud As Long

    Dim FLAG     As Long

    Dim found    As Integer

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
    ybuf = menup.vservidor

    If IsValidIPAddress(menup.vservidor) = False Then
        'ybuf = Trim(RecuperarIP)
        'ybuf = "" & menup.Winsock1.LocalIP
        'MsgBox ybuf
        ybuf = GetMACs_AdaptInfo()

    End If

    'MsgBox ybuf
    'If Len(menup.vservidor) > 0 Then
    If GetRemoteMACAddress(ybuf, xbuf, "") Then
        'MsgBox xbuf
        xbuf = xbuf
    Else
        xbuf = ""

    End If

    'MsgBox xbuf
    If Trim(Len(xbuf)) = 0 Then
        xbuf = GetMACs_AdaptInfo()

    End If

    'MsgBox xbuf
      
    'End If
    'End If
    'MsgBox xbuf
    '-----solo le puse por el momento----
    'xbuf = ""
    
    unidad = "C:\"
    Call GetVolumeInformation(unidad, cad1, 256, numSerie, longitud, FLAG, cad2, 256)
    buf = xbuf + "" & numSerie
    serie_disco_duro = buf
    'MsgBox "paso"
    Exit Function
cmd9011_err:
    MsgBox "Aviso en shd" + error$, 48, "Aviso"
    Exit Function

End Function

Function serie_mac(abuf As String) As String

    Dim unidad As String

    Dim buf    As String

    Dim xbuf   As String

    Dim ybuf   As String

    On Error GoTo cmd19011_err

    Dim cad1     As String * 256

    Dim cad2     As String * 256

    Dim numSerie As Long

    Dim longitud As Long

    Dim FLAG     As Long

    Dim found    As Integer

    '
    
    'anterior fueron estos dos
    'unidad = "C:\"
    'Call GetVolumeInformation(unidad, cad1, 256, numSerie, longitud, flag, cad2, 256)
    'buf = numSerie
    'fin a de anterior
    
    'ahora haciendo con la serie disco duro
    'serie_disco_duro = placa_madre()
    'MsgBox "Numero de Serie de la unidad " & unidad & " = " & numSerie
    'MsgBox abuf
    xbuf = ""
    ybuf = menup.vservidor
    ybuf = abuf

    'If Len(menup.vservidor) > 0 Then
    'xbuf = CpuId()
    'MsgBox xbuf
    If GetRemoteMACAddress(ybuf, xbuf, "") Then
        xbuf = xbuf
    Else
        xbuf = ""

    End If

    'MsgBox xbuf
       
    '   End If
    'End If
    'MsgBox xbuf
    'xbuf = ""
    
    'unidad = "C:\"
    'Call GetVolumeInformation(unidad, cad1, 256, numSerie, longitud, flag, cad2, 256)
    'buf = xbuf + "" & numSerie
    
    'MsgBox xbuf
    serie_mac = xbuf
    'MsgBox "paso"
    Exit Function
cmd19011_err:
    MsgBox "Aviso en shd " + error$, 48, "Aviso"
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

    Dim I As Integer

    On Error Resume Next

    I = Len(Dir$(FileName))

    If Err Or I = 0 Then
        bFileExists = False
    Else
        bFileExists = True

    End If

End Function

Function existearchivo(FileName As String) As Integer

    Dim I As Integer

    On Error Resume Next

    I = Len(Dir$(FileName))

    If Err Or I = 0 Then
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

Function copiar_servicio()

    On Error GoTo cmd239_err

    FileCopy globaldat & "\tdetalle.dbf", globaldat & "\" & "_s" & gusuario & ".dbf"
    FileCopy globaldat & "\tdetalle.cdx", globaldat & "\" & "_s" & gusuario & ".cdx"
    copiar_servicio = 1
    Exit Function
cmd239_err:
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

Function proceso_formatos(archivo_formato As String, _
                          mytablex As ADODB.Recordset, _
                          ubicacioni As String, _
                          ubicacionf As String, _
                          basedatos As String, _
                          indice As String, _
                          bxlocal As String, _
                          tipo As String, _
                          bxserie As String, _
                          Numero As String, _
                          ascopia As String, _
                          contando As Integer)

    ',nombre as string)
    On Error GoTo cmd56789_err

    Dim linea$

    Dim buff$

    Dim campo       As String

    Dim j           As Integer

    Dim sw          As Integer

    Dim posicioni   As Long

    Dim posicionf   As Long

    Dim tlinea      As String

    Dim valor       As String

    Dim found       As Integer

    Dim nombrearch  As String

    Dim nombrearch1 As String

    Dim posicionb   As Long

    Dim variable    As String

    Dim sw1         As Integer

    Dim bufx        As String

    Dim xxsw        As Integer

    Dim alibaba     As Integer

    Dim antfont

    'inicio 30/05/2017 pll
    Dim mynombre

    Dim posini As Integer

    Dim Texto  As String

    posini = 0
    Texto = ""
    'fin 30/05/2017 pll

    cerrar_archivo
    antfont = Printer.FontSize
                   
    'MsgBox archivo_formato
    nombrearch = globaldir & "\temporal\" & gusuario & ".txt"
    nombrearch1 = globaldir & "\formatos\" & archivo_formato
    'MsgBox nombrearch & " " & nombrearch1
    
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
        'inicio 05/02/2018 pll
        'linea = Mid$(buff, 1, Len(buff))
        linea = Mid$(Replace(buff, "Ñ", "N"), 1, Len(buff))
        linea = Mid$(Replace(buff, "&", "Y"), 1, Len(buff))
        'fin 05/02/2018 pll
          
        'inicio 29/05/2017 pll
        'If linea = "CLIENTE:[NOMBRE,1,80,]" Then
           
        '
        '           If nombre > "" Then
        '           MsgBox (nombre)
        '            If InStr(linea, "[NOMBRE,1,") <> 0 Then
        '             mynombre = Len(nombre)
        '            '  MsgBox (mynombre)
        '             If mynombre >= 81 Then
        '                mynombre = 80
        '             End If
        '
        '             posini = InStr(linea, "[")
        '             Texto = Mid$(linea, 1, posini - 1)
        '             linea = Texto + "[NOMBRE,1," & mynombre & ",]"
        '
        '          End If
        '          'fin 29/05/2017 pll
          
        '  End If
          
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
                    'If variable <> "@" And variable <> "+" Then
                    found = formateaa(variable, 1, 0, 0)

                    'End If
                End If

                'If Mid$(linea, j, 1) = "@" Then  'negrita
                '   antfont = Printer.FontSize
                '   Printer.FontSize = 13
                '   MsgBox Printer.FontSize
                '   found = formateaa("Hola", 10, 2, 0)
                '   'Printer.FontSize = antfont
                'End If
                'If Mid$(linea, j, 1) = "+" Then  'negrita
                '   MsgBox antfont
                '   Printer.FontSize = antfont
                '   found = formateaa("Hola", 10, 2, 0)
                '   Printer.FontSize = antfont
                'End If
                If Mid$(linea, j, 1) = "[" Then
                    sw = 1
                    posicioni = j + 1

                End If

                If sw = 1 And Mid$(linea, j, 1) = "]" Then
                    posicionf = j - 1
                    campo = Mid$(linea, posicioni, posicionf - posicioni + 1)
                    alibaba = 0
                    valor = busca_campo1(basedatos, mytablex, campo, bxlocal, bxserie, Numero, contando, alibaba, ascopia)
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
    MsgBox "xx.Existe Un error en Proceso Formatos-- " & error$, 24, "Aviso"
    cerrar_archivo
    Exit Function
error_lectura:
    MsgBox "Error en Proceso_formatos", 24, "Aviso"
    cerrar_archivo
    Exit Function
    
End Function

Function numero_diasMes()

    Dim month_number As Integer

    Dim year_number  As Integer

    'month_number = Month(txtMonth.Text)
    'year_number = Year(txtMonth.Text)
    'MsgBox "Days: " & Format$(Day(DateSerial(year_number, month_number + 1, 0)))
End Function

Function busca_campo1(tablabasedatos As String, _
                      mytablex As ADODB.Recordset, _
                      campo As String, _
                      bxlocal As String, _
                      bxserie As String, _
                      tablas As String, _
                      contando As Integer, _
                      alibaba As Integer, _
                      ascopia As String) As String

    Dim knik1      As String

    Dim knik2      As String

    Dim knik11     As String

    Dim knik22     As String

    Dim amigohs    As String

    Dim CAMPO1     As String

    Dim CAMPO2     As String

    Dim campo3     As String

    Dim campo4     As Integer

    Dim found      As Integer

    Dim sdx        As Double

    Dim campoz     As String

    Dim campoy     As String

    Dim ponemoneda As String

    Dim buf        As String

    Dim sdx1       As Double

    Dim j          As Integer

    Dim mytabley   As New ADODB.Recordset

    'Dim mydby As Database
    Dim ddd        As String

    Dim mmm        As String

    Dim yyy        As String

    Dim bufx       As String

    Dim xtamano    As String

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
        If campoz = "DETALLE" Then
            mytabley.Open "select * from detalle where local='" & "" & mytablex.Fields("local") & "' and tipo='" & "" & mytablex.Fields("tipo") & "' and serie='" & "" & mytablex.Fields("serie") & "' and numero='" & "" & mytablex.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

            If mytabley.RecordCount > 0 Then 'si existe
                'MsgBox CAMPO1
                '----------------------------
                found = extraer_campos(campoy, CAMPO1, CAMPO2, campo3, campo4, ",")
                buf = Trim("" & mytabley.Fields(CAMPO1))

                'found = formateaa(buf, Val(campo3), 0, 0)
                Select Case campo4

                    Case 4
                        found = formateaa(buf, Val(campo3), 0, 1)

                    Case Else
                        found = formateaa(buf, Val(campo3), 0, 0)

                End Select

                '----------------------------
            End If

            mytabley.Close
            Exit Function
       
        End If

        If campoz = "PRODUCTO" Then
            mytabley.Open "SELECT * FROM producto where  producto='" & "" & mytablex.Fields("producto") & "'", cn, adOpenDynamic, adLockOptimistic

            If mytabley.RecordCount > 0 Then 'si existe
                'MsgBox CAMPO1
                '----------------------------
                'inicio 05/02/2018 pll
                'buf = Trim("" & mytabley.Fields(CAMPO1))
                found = extraer_campos(campoy, CAMPO1, CAMPO2, campo3, campo4, ",")
                buf = Replace(Trim("" & mytabley.Fields(CAMPO1)), "Ñ", "N")

                'fin 05/02/2018 pll
                'found = formateaa(buf, Val(campo3), 0, 0)
                Select Case campo4

                    Case 4
                        found = formateaa(buf, Val(campo3), 0, 1)

                    Case Else
                        found = formateaa(buf, Val(campo3), 0, 0)

                End Select

                '----------------------------
            End If

            mytabley.Close
            Exit Function

        End If

        If campoz = "PRECIOS" Then
            'MsgBox "" & mytablex.Fields("producto")
            mytabley.Open "SELECT * FROM precios where  producto='" & "" & mytablex.Fields("producto") & "' and local='" & "" & mytablex.Fields("local") & "'", cn, adOpenDynamic, adLockOptimistic

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
                'inicio 06/02/2018 pll
                'buf = Trim("" & mytabley.Fields(CAMPO1))
                buf = Replace(Trim("" & mytabley.Fields(CAMPO1)), "Ñ", "N")
                'fin 06/02/2018 pll
                'If Val(campo3) > Val(CAMPO2) Then
                'buf = Mid$(Trim("" & mytabley.Fields(CAMPO1)), Val(CAMPO2), Val(campo3))
                'End If
                'MsgBox CAMPO2 & " " & campo3
                'MsgBox CAMPO1 & "" & CAMPO2 & "" & campo3 & "" & campo4
                'MsgBox Mid$(buf, Val(campo3), Val(campo4))
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

        If campoz = "CLASESUNAT" Then
            'Set mydby = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
            'If Len(Trim("" & mytablex.Fields("clasesunat"))) > 0 Then
            mytabley.Open "SELECT * FROM clasesunat where  clasesunat='" & Trim("" & mytablex.Fields("clasesunat")) & "'", cn, adOpenDynamic, adLockOptimistic

            If mytabley.RecordCount > 0 Then 'si existe
                '----------------------------
                found = extraer_campos(campoy, CAMPO1, CAMPO2, campo3, campo4, ",")
                buf = "" & mytabley.Fields(CAMPO1)
                found = formateaa(buf, Val(campo3), 0, 0)
                '----------------------------
            Else
                buf = "PERCEPCION"
                found = extraer_campos(campoy, CAMPO1, CAMPO2, campo3, campo4, ",")
                'MsgBox campo3
                found = formateaa(buf, Val(campo3), 0, 0)

            End If

            mytabley.Close
            'End If
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
            Exit Function

        End If

        If campoz = "TLOCAL" Then
            'Set mydby = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
            mytabley.Open "SELECT * FROM tlocal where  codigo='" & "" & bxlocal & "'", cn, adOpenDynamic, adLockOptimistic

            If mytabley.RecordCount > 0 Then 'si existe
                '----------------------------
                'MsgBox "abc"
                found = extraer_campos(campoy, CAMPO1, CAMPO2, campo3, campo4, ",")
                buf = Trim("" & mytabley.Fields(CAMPO1))
                found = formateaa(buf, Val(campo3), 0, 0)

                '----------------------------
            End If

            mytabley.Close
            Exit Function

        End If
   
        If campoz = "TIPO" Then
            'Set mydby = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
            mytabley.Open "SELECT * FROM TIPO where  codigo='" & "" & mytablex.Fields("TIPO") & "'", cn, adOpenDynamic, adLockOptimistic

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
            'MsgBox "xxx"
            'Set mydby = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
            mytabley.Open "SELECT * FROM " & campoz & " where  local='" & "" & mytablex.Fields("local") & "' and tipo='" & "" & mytablex.Fields("tipo") & "' and serie='" & "" & mytablex.Fields("serie") & "' and numero='" & "" & mytablex.Fields("numero") & "'", cn, adOpenDynamic, adLockOptimistic

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

        If campoz = "SALDO" Then
            sdx = Val("" & mytablex.Fields("total")) - Val("" & mytablex.Fields("acuenta"))
            buf = Format(sdx, "0.00")
            found = extraer_campos(campoy, CAMPO1, CAMPO2, campo3, campo4, ",")
            found = formateaa(buf, Val(campo3), 0, 0)
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
        If Mid$(campo, 1, 1) = "+" Then   'si percepcion
            CAMPO1 = Mid$(CAMPO1, 2, Len(CAMPO1) - 1)
            buf = ""

            If Trim("" & mytablex.Fields(CAMPO1)) = "S" Then
                buf = "P"

            End If

            buf = Mid$(buf, Val(CAMPO2), Val(campo3))
            found = formateaa(buf, Val(campo3), 0, 0)
            Exit Function

        End If

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

            If "" & mytablex.Fields("servicio") = "A" Then
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
            found = imprime_delivery(mytablex)
            Exit Function

        End If
   
        If UCase$(campo) = "PONEMONEDA" Then
            ponemoneda = signopeso

            If "" & mytablex.Fields("moneda") = "S" Then
                ponemoneda = dicmoneda

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
                ponemoneda = dicmoneda

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

    'inicio 04/12/2017 pll
    If UCase$(CAMPO1) = "CDR" Then
        If Len("" & mytablex.Fields("cdr")) > 0 Then
            'buf = Space$(10) & "" & mytablex.Fields("cdr") 'aqui es para convertir qr
            buf = "" & mytablex.Fields("cdr")
            found = formateaa(buf, Val(campo3), 0, 0)

        End If

        Exit Function

    End If

    'fin 04/12/2017 pll
    'inicio 27/12/29017 pll 'esto es para la nota debito
    If UCase$(CAMPO1) = "CDR_NDV" Then
        If Len("" & mytablex.Fields("cdr_ndv")) > 0 Then
            buf = "" & mytablex.Fields("cdr_ndv")
            found = formateaa(buf, Val(campo3), 0, 0)

        End If

        Exit Function

    End If

    'aqui nota credito
    If UCase$(CAMPO1) = "CDR_NCV" Then
        If Len("" & mytablex.Fields("cdr_ncv")) > 0 Then
            buf = "" & mytablex.Fields("cdr_ncv")
            found = formateaa(buf, Val(campo3), 0, 0)

        End If

        Exit Function

    End If

    'fin 27/12/2017 pll
    If UCase$(campo) = "DELIVERY" Then
        found = imprime_delivery(mytablex)
        Exit Function

    End If

    '----------------------- -----------------

    If Val(CAMPO2) > 0 And Val(campo3) > 0 Then
        If CAMPO1 <> "DOCUMENTO" Then
            If UCase$(CAMPO1) = "NOMBRE" Then
                buf = Replace(Trim("" & mytablex.Fields(CAMPO1)), "Ñ", "N")
            Else
                buf = Mid$("" & mytablex.Fields(CAMPO1), Val(CAMPO2), Val(campo3))

            End If

        End If
    
    Else
        buf = "" & mytablex.Fields(CAMPO1)

    End If

    '30/10/2017 Impresión de Tipo de documento en comprobantes...Segunda Opcion

    '30/10/2017 Impresión de Tipo de documento en comprobantes...Segunda Opcion

    'If UCase$(CAMPO1) = "REFERENCIA" Then
    '   buf = busca_cajay("" & "" & mytablex.Fields("caja"))
    '
    'End If
    'si es campo5 otra forma --------- tipo de letra imprimir---

    '-------------------------------------------------------------
    'MsgBox tablabasedatos

    '''30/10/2017 Impresión de Tipo de documento en comprobantes
    '''30/10/2017 Impresión de Tipo de documento en comprobantes

    If CAMPO1 <> "DOCUMENTO" And CAMPO1 <> "TIPO1" And CAMPO1 <> "SERIE1" And CAMPO1 <> "NUMERO1" Then

        Select Case Val("" & mydbxglo.TableDefs(tablabasedatos).Fields(CAMPO1).Type)

                'Select Case Val("" & mytablex.Fields(CAMPO1).Type)
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

    End If

    '''30/10/2017 Impresión de Tipo de documento en comprobantes
    If CAMPO1 = "DOCUMENTO" Then
        buf = "" & busca_tipocomprobante(mytablex.Fields("tipo1"))
        found = formateaa(buf, Val(campo3), 0, 0)

    End If
     
    If CAMPO1 = "TIPO1" Then
        buf = "" & mytablex.Fields("tipo1")
        found = formateaa(buf, Val(campo3), 0, 0)

    End If

    If CAMPO1 = "SERIE1" Then
        buf = "" & mytablex.Fields("serie1")
        found = formateaa(buf, Val(campo3), 0, 0)

    End If

    If CAMPO1 = "NUMERO1" Then
        buf = "" & mytablex.Fields("numero1")
        found = formateaa(buf, Val(campo3), 0, 0)

    End If

    '''30/10/2017 Impresión de Tipo de documento en comprobantes

    Exit Function
cmd9876_err:
    MsgBox "Error en busca Campo1 " & campo & error$, 24, "Aviso"
    Exit Function

End Function

''30/10/2017 Impresión de Tipo de documento en comprobantes
Function busca_tipocomprobante(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT descripcio FROM tipo where  tipo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        busca_tipocomprobante = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close

End Function

''30/10/2017 Impresión de Tipo de documento en comprobantes

Sub cerrar_archivo()

    On Error GoTo cmd561_err

    Close
    Exit Sub
cmd561_err:
    MsgBox "Aviso en cerrar_archivo " + error$, 48, "Aviso"
    Exit Sub

End Sub

Function extraer_campos(campo As String, _
                        CAMPO1 As String, _
                        CAMPO2 As String, _
                        campo3 As String, _
                        campo4 As Integer, _
                        Flags As String)

    Dim I    As Integer

    Dim j    As Integer

    Dim temp As String

    I = 0
    temp = Trim$(campo)

    If Len(temp) = 0 Then Exit Function
    Do
        j = InStr(temp, Flags)

        If j > 0 Then
            I = I + 1

            Select Case I

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

Function extraer_campos1(campo As String, _
                         CAMPO1 As String, _
                         CAMPO2 As String, _
                         campo3 As String)

    Dim I    As Integer

    Dim j    As Integer

    Dim temp As String

    I = 0
    temp = Trim$(campo)

    If Len(temp) = 0 Then Exit Function
    Do
        j = InStr(temp, "|")

        If j > 0 Then
            I = I + 1

            Select Case I

                Case 1: CAMPO1 = Mid$(temp, 1, j - 1)

                Case 2: CAMPO2 = Mid$(temp, 1, j - 1)

                Case 3: campo3 = Mid$(temp, 1, j - 1)

                    'Case 4: campo4 = CInt(Mid$(temp, 1, j - 1))
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

    Dim sdx   As Double

    Dim buf   As String

    Dim buf1  As String

    Dim buf2  As String

    Dim found As Integer

    Dim ik    As Integer

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
        If Trim(menup.Label10) = "ARGENTINA" Then
            buf2 = buf2 & "/100 PESOS "

        End If

        If Trim(menup.Label10) = "PERU" Then
            buf2 = buf2 & "/100  SOLES"

        End If
       
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

Function busca_serialesss(aatipo As String, _
                          aanumero As String, _
                          aaproducto As String, _
                          contando As Integer, _
                          CAMPO2 As Integer, _
                          campo3 As Integer)

    '
End Function

Function busca_tallasx(mytablex, _
                       contando As Integer, _
                       CAMPO2 As Integer, _
                       campo3 As Integer)

    Dim buf         As String

    Dim mytabley    As New ADODB.Recordset

    Dim xtallas(17) As String

    Dim ytallas(17) As Double

    Dim I           As Integer

    Dim found       As Integer

    On Error GoTo cmd451213_err

    mytabley.Open "select * from linea where linea='" & "" & mytablex.Fields("linea") & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount > 0 Then

        For I = 1 To 16
            xtallas(I) = "" & mytabley.Fields("t" & I)
        Next I

    End If

    mytabley.Close

    For I = 1 To 16
        ytallas(I) = Val("" & mytablex.Fields("t" & I))
    Next I

    buf = ""

    For I = 1 To 16

        If ytallas(I) > 0 Then
            buf = buf & xtallas(I) & "/" & ytallas(I) & " "

        End If

    Next I

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

    Dim temp  As String

    Dim j     As Integer

    Dim Flags As String

    Dim buf1  As String

    Dim sw    As Integer

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

Function verifica_receta_flag1(buf As String, _
                               contando As Integer, _
                               CAMPO2 As Integer, _
                               campo3 As Integer)

    Dim temp  As String

    Dim j     As Integer

    Dim Flags As String

    Dim buf1  As String

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

    Dim inga     As Long

    Dim negativo As Variant

    Dim L        As Integer

    Dim una      As Variant

    Dim millon   As Variant

    Dim millones As Variant

    Dim vez      As Integer

    Dim maxvez   As Integer

    Dim k        As Integer

    Dim strq     As String

    Dim strb     As String

    Dim stru     As String

    Dim strd     As String

    Dim strc     As String

    Dim ia       As Integer

    Dim strn()   As String

    Dim lo       As Integer

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

Sub factura_formato(bxlocal As String, _
                    bxtipo As String, _
                    bxserie As String, _
                    bxnumero As String, _
                    ascopia As String, _
                    sw As Integer)

    Dim mytablex        As New ADODB.Recordset

    Dim mytabley        As New ADODB.Recordset

    Dim mytablez        As New ADODB.Recordset

    Dim vacu            As String

    Dim buf             As String

    Dim found           As Integer

    Dim nro_lineas      As Integer

    Dim contando        As Integer

    Dim faltante        As Integer

    Dim I               As Integer

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

    If sw = 0 Then  'si es 1 no borra archivo formato
        found = borra_nombre("" & FileName)

    End If

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
             
    'inicio 30/05/2017 pll para que viaje la variable nombre
    'found = proceso_formatos(archivo_formato, mytabley, "{", "}", cgusuario, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    found = proceso_formatos(archivo_formato, mytabley, "{", "}", cgusuario, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    'fin 30/05/2017 pll para que viaje la variable nombre
       
    vacu = "" & mytabley.Fields("acu")
    'MsgBox ""
    '
    'detalle
    flag_contando = 0
    mytablex.Open "select * from " & dgusuariog & " where local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        Do

            If mytablex.EOF Then Exit Do
            If "" & mytablex.Fields("dua") <> "R" Then
                flag_contando = contando + 1
             
                'inicio 30/05/2017 pll para el nombre del cliente de la tickera
                'found = proceso_formatos(archivo_formato, mytablex, "/", "\", dgusuariog, "TDETALLE", bxlocal, bxtipo, bxserie, bxnumero, ascopia, contando)
                found = proceso_formatos(archivo_formato, mytablex, "/", "\", dgusuariog, "TDETALLE", bxlocal, bxtipo, bxserie, bxnumero, ascopia, contando)
                'fin 30/05/2017 pll para el nombre del cliente de la tickera
             
                contando = contando + 1

            End If

            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
        
    If bxtipo <> "1" And bxtipo <> "2" And bxtipo <> "5" And bxtipo <> "9" Then

        'If vacu = "V" Or vacu = "F" Or vacu = "P" Or vacu = "D" Or vacu = "I" Or vacu = "S" Then
        'If vacu = "V" Or vacu = "F" Or vacu = "P" Or vacu = "D" Or vacu = "I" Then
        If contando < nro_lineas Then

            For I = contando To nro_lineas
                Open FileName For Append As #1
                found = formateaa("", 1, 2, 0)
                Close #1
            Next I

            'End If
        End If

    End If

    'total  xxxx
       
    'Set mytablex = mydbxglo.OpenTable(cgusuario)
    'mytablex.Index = "tfactura"
       
    'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    'found = proceso_formatos(archivo_formato, mytabley, "$", "?", cgusuario, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    found = proceso_formatos(archivo_formato, mytabley, "$", "?", cgusuario, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera       mytabley.Close
    
    If Len(gofpago) > 0 Then
        mytablez.Open "SELECT * FROM " & gofpago & "   where  local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic

        If mytablez.RecordCount > 0 Then 'si existe
            Do

                If mytablez.EOF Then Exit Do
                'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
                'found = proceso_formatos(archivo_formato, mytablez, "<", ">", gofpago, "fpagov", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
                found = proceso_formatos(archivo_formato, mytablez, "<", ">", gofpago, "fpagov", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
                'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
                mytablez.MoveNext
            Loop

        End If

        mytablez.Close

    End If
       
    If sw = 1 Then
        Open FileName For Append As #1
        buf = Chr$(27) + "i"  'epson
        Print #1, buf;
        Close #1

    End If
       
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

Function imprime_archivoj(xbuf As String, xsw As Integer, xtipoletra As String)

    Dim I         As Long

    Dim max       As Long

    Dim buf       As String

    Dim free_file As Integer

    Dim vr

    Dim antfont

    On Error GoTo cmd9876_err

    free_file = FreeFile
    Open xbuf For Input As free_file
    max = LOF(free_file)   'numero de letras

    If max < 1 Then
        Close free_file
        Exit Function

    End If

    ponerfont "Courier New"

    'Printer.FontName = "Courier New"
    If Val(xtipoletra) < 7 Then
        xtipoletra = "9"

    End If

    Printer.FontBold = False
    Printer.FontSize = Val(xtipoletra)

    For I = 1 To max
        vr = DoEvents()

        If opcion3 = "1" Then
            If MsgBox("DESEA CANCELAR LA IMPRESION", 1, "AVISO") = 1 Then
                Close free_file
                Printer.EndDoc
                Exit Function

            End If

            opcion3 = "0"

        End If

        Seek free_file, I
        buf = input$(1, free_file)
        'MsgBox buf
        ' If buf = "@" Then
        '   MsgBox "HOLA"
        '   antfont = Printer.FontSize
        '   Printer.FontSize = 13
        'End If
        'Printer.Print Palabra;
        'If buf = "+" Then
        '   Printer.FontSize = antfont
        'End If
        
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
    Next I

    Close free_file
    Printer.Print
    Printer.EndDoc
    Exit Function
cmd9876_err:
    MsgBox "Error en imprime archivo j " & error$, 48, "Aviso"
    Printer.EndDoc
    Exit Function

End Function

Function imprime_archivotexto(path As String)

    Dim txtTheLine As String

    Printer.FontName = "Courier New"
    Printer.FontSize = 9
    Printer.FontBold = False
    Printer.Print ""
    Open path For Input As 1

    Do While Not EOF(1)
        Line Input #1, txtTheLine '& VarTexto & vbCrLf   'Read a line
        Printer.Print txtTheLine      'send to the default printer.
    Loop
    Printer.Print ""
    Printer.EndDoc  'finally stop printing after a Form Feed.
    Close #1

End Function

Sub ponerfont(nombrefont As String)

    On Error GoTo cm78999_err

    If Len(Trim(nombrefont)) > 0 Then
        Printer.FontName = Trim(nombrefont)

        'MsgBox "abc"
        'Printer.FontName = Trim(nombrefont)
    End If

    Exit Sub
cm78999_err:
    MsgBox "Fuente Inpresora no existe ", 48, "Aviso"
    Exit Sub

End Sub

Function busca_combox(xxyzcontrol As Control, buf As String)

    On Error GoTo cmd45_err

    Dim I  As Integer

    Dim sw As Integer

    sw = 0

    For I = 0 To xxyzcontrol.ListCount - 1

        If xxyzcontrol.List(I) = buf Then
            busca_combox = I
            sw = 1
            Exit For

        End If

    Next I

    If sw = 0 Then
        busca_combox = 0

    End If

    Exit Function
cmd45_err:
    busca_combox = 0
    Exit Function

End Function

Function calcula_saldo(sdx As Double, sdx1 As Double) As String

    Dim buf    As String

    Dim dsdx   As Double

    Dim rsdx   As Double

    Dim signo1 As Variant

    Dim buf1   As String

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

    Dim buf  As String

    buf = Format(sdx, "#.#")

    If InStr(buf, ".") Then
        extraer_decimal = Val(Mid$(buf, InStr(buf, ".") + 1, 2))
    Else
        extraer_decimal = 0

    End If

End Function

Function etiqueta_disco() As String

    Dim unidad   As String

    Dim cad1     As String * 256

    Dim cad2     As String * 256

    Dim numSerie As Long

    Dim longitud As Long

    Dim FLAG     As Long

    unidad = "c:\"
    Call GetVolumeInformation(unidad, cad1, 256, numSerie, longitud, FLAG, cad2, 256)
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

Function extra_loquesea1(buf As String) As String

    Dim j

    Dim buf1 As String

    buf1 = ""

    If InStr(buf, "|") > 0 Then
        j = InStr(buf, "|")
        buf1 = Mid$(buf, j + 1, Len(buf) - (j))
    Else
        buf1 = buf

    End If

    extra_loquesea1 = buf1

End Function

Sub cabecera_tipico(buf1 As String, buf2 As String, buf3 As String)

    Dim buf      As String

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset
   
    buf = ""
   
    mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        buf = buf & "" & mytablex.Fields("cabecera1")

    End If

    mytablex.Close
   
    '''24/08/2017  Kenyo descripcion larga en reportes ticket
    'found = formateaa(buf, 70, 0, 0)
  
    If repinv.gcanti = "S" Then
        found = formateaa(buf, 52, 0, 0)
    Else
        found = formateaa(buf, 62, 0, 0)

    End If
   
    '''24/08/2017  Kenyo descripcion larga en reportes ticket
   
    buf = formateaa(" ", 5, 0, 0)
    buf = "Pagina :" & contpag
    found = formateaa(buf, 15, 2, 0)
   
    buf = "Usuario :" & buf3
   
    '''24/08/2017  Kenyo descripcion larga en reportes ticket
    'found = formateaa(buf, 70, 0, 0)
    If repinv.gcanti = "S" Then
        found = formateaa(buf, 52, 0, 0)
    Else
        found = formateaa(buf, 62, 0, 0)

    End If

    '''24/08/2017  Kenyo descripcion larga en reportes ticket
   
    buf = formateaa(" ", 5, 0, 0)
    buf = "Hora   :" & Format(Now, "hh:mm:ss")
   
    found = formateaa(buf, 20, 2, 0)
    buf = formateaa("", 1, 2, 0)

End Sub

''' /04/2017 kenyo. Formato ticket de cuentas por cobrar y pagar en formato ticket
Sub cabecera_tipico2(buf1 As String, buf2 As String, buf3 As String)

    Dim buf      As String

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset
   
    buf = ""
    mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        buf = buf & "" & mytablex.Fields("cabecera1")

    End If

    mytablex.Close
    found = formateaa(buf, 70, 0, 0)
    buf = formateaa("", 1, 2, 0)
    buf = "Hora   :" & Format(Now, "hh:mm:ss")
    found = formateaa(buf, 20, 2, 0)
    buf = formateaa("", 1, 2, 0)

End Sub

''' /04/2017 kenyo. Formato ticket de cuentas por cobrar y pagar en formato ticket

Function crypt(pW, cryptee)

    Dim I

    Dim pchar

    Dim cchar

    Do While Len(pW) < Len(cryptee)
        pW = pW & pW
    Loop

    For I = 1 To Len(cryptee)
        pchar = Asc(Mid$(pW, I, 1))
        cchar = Asc(Mid$(cryptee, I, 1))
        Mid$(cryptee, I, 1) = Chr$(pchar Xor cchar)
    Next I

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

Function ejecuta_wordpad(buf As String)

    Dim mytablex As New ADODB.Recordset

    Dim buf1     As String

    buf1 = ""
    'MsgBox buf
    mytablex.Open "Select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        buf1 = Trim("" & mytablex.Fields("wordpad"))

    End If

    mytablex.Close

    If Len(buf1) > 0 Then
        ExecuteCommand buf1 & " " & buf

    End If

    ejecuta_wordpad = 1

End Function

Function valida_wordpad(buf As String)

    Dim mytablex As New ADODB.Recordset

    Dim buf1     As String

    buf1 = ""
    'MsgBox buf
    ejecutawor = 0
    mytablex.Open "Select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        buf1 = Trim("" & mytablex.Fields("wordpad"))

    End If

    mytablex.Close

    genver.file = buf

    If VISTA <> "OK" Then
        genver.Show 1

    End If

    '
   
    'End If
    If ejecutawor = 1 Then
        If Len(buf1) > 0 Then
            ExecuteCommand buf1 & " " & buf

        End If

    End If

    valida_wordpad = 1
    
End Function

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

    Dim X         As Integer

    Dim strError  As String

    Dim lngStatus As Integer

    On Error GoTo cmd8911_err

    Select Case slpt

        Case "COM1"
            X = 1

        Case "COM2"
            X = 2

        Case "COM3"
            X = 3

        Case "COM4"
            X = 4

        Case "COM5"
            X = 5

    End Select
       
    If inicializa_mscomm(X) <> 1 Then
        Exit Function

    End If

    If escribe_mscomm(b) <> Len(b) Then

    End If

    cerrar_mscomm
    imprime_linea = 1
       
    Exit Function
cmd8911_err:
    MsgBox "Aviso en imprime linea ", 48, "Aviso"
    Exit Function

End Function

Function imprime_puerto_serial(slpt As String)

    On Error GoTo cmd4312_err

    Dim jx        As Integer

    Dim found     As Integer

    Dim X         As Integer

    Dim I         As Integer

    Dim max       As Long

    Dim buf       As String

    Dim strError  As String

    Dim lngStatus As Integer

    Select Case slpt

        Case "COM1"
            X = 1

        Case "COM2"
            X = 2

        Case "COM3"
            X = 3

        Case "COM4"
            X = 4

        Case "COM5"
            X = 5

    End Select
           
    'FlushComm
    'MsgBox X
    If inicializa_mscomm(X) <> 1 Then
        Exit Function

    End If
       
    found = SendFile("" & FileName)
       
    'Do While Not EOF(1)
    '    Line Input #1, buf
    '    If escribe_mscomm(buf) <> Len(buf) Then
    '    End If
    'Loop
       
    'jx = 0
    'max = LOF(1)   'numero de letras
    'For i = 1 To max
    '    Seek #1, i
    '    buf = Input$(1, #1)
    '    If escribe_mscomm(buf) <> Len(buf) Then
    '    End If
    '    'jx = jx + 1
    '    'If jx > 50 Then
    '    '   'MsgBox "xxx"
    '    '   Sleep (1)
    '    '   jx = 0
    '    'End If
    'Next i
    'MsgBox "xx"
    'Close #1
    'Sleep (1)

    cerrar_mscomm
    cerrar_archivo
    Exit Function
cmd4312_err:
    MsgBox "Aviso en imprime_puerto_serial " + error$, 48, "Aviso"
    cerrar_archivo
    Exit Function

End Function

Function SendFile(Tmp$)

    Dim I

    Dim tiempo

    Dim ij As Integer

    Dim temp$

    Dim hsend, bsize, LF&

    Dim ret

    Dim vr

    Dim SendLen As Integer

    I = FreeFile
    Open Tmp$ For Binary Access Read As #I
    bsize = menup.MSComm1.OutBufferSize
    LF& = LOF(I)

    Do Until EOF(I)

        If LF& - Loc(I) <= bsize Then
            bsize = LF& - Loc(I) + 1

        End If

        temp$ = Space$(bsize)
        Get #I, , temp$
        '----------------------
        SendLen = Len(temp$)

        For ij = 1 To SendLen
            'menup.MSComm1.InBufferCount = 0
            menup.MSComm1.Output = Mid$(temp$, ij, 1)
            vr = DoEvents()
            'espera_segundo 1000
            'Sleep (0.2)
            'menup.MSComm1.Output = Len(Mid$(temp$, ij, 1))
            'MsgBox menup.MSComm1.OutBufferCount
            'WaitMS 0.1
            'Sleep (0.9)
            'menup.MSComm1.Output = temp$
      
            'Do
            'vr = DoEvents()
            'Loop Until menup.MSComm1.OutBufferCount = 0
            'vr = DoEvents()
            'tiempo = Now
            'If menup.MSComm1.OutBufferCount > 0 Then
            ' Do
            '     vr = DoEvents()
            '     If DateDiff("s", Now, tiempo) > 10 Then
            '        If MsgBox("Datos No enviados", 1, "Desea Reintentar..") <> 1 Then Exit Do
            '     End If
            'Loop
            'End If
        Next ij

        'Sleep (50)
    
    Loop
    'Sleep (20)
    'MsgBox "Presione Enter para Continuar...", 48, "Aviso"
    Close #I

End Function

Function star_sp342(Puerto As String, sw As Integer) 'sw determina tipo de dispositivo

    Dim R%

    Dim s         As Boolean

    Dim found     As Integer

    Dim buf       As String

    Dim vr        As Integer

    Dim bufferr   As String

    Dim slpt      As String

    Dim I         As Long

    Dim j         As Integer

    Dim max       As Long

    Dim velox     As String

    Dim lngStatus As Long

    Dim X         As Integer

    Dim strError  As String

    Dim nficsal   As Integer

    Dim b         As String

    Dim contin    As Integer

    On Error GoTo cmd13081_err

    If Len(FileName) = 0 Then
        MsgBox "Mensaje,Nombre archivo no existe ", 24, "Aviso"
        Exit Function

    End If

    I = 0
    cerrar_archivo
   
    slpt = Trim(Puerto)
   
    ''11/07/2017 kenyo multicomandas
      
    '   If puerto = "BAR" Then
    '      'impresion_codbar "IMPRIME"
    '      Exit Function
    '   End If
    
    '' 29/11/2017 Correcion de bloqueo de botones  al registrar comandas
    If Puerto = "" Or Puerto = " " Then
        Puerto = "%"

    End If

    '' 29/11/2017 Correcion de bloqueo de botones  al registrar comandas
    
    If nroimpresion = 0 Then ' puertoimpresion
        If Puerto = "%" Then
            Exit Function

        End If

    End If
    
    ' PUERTOIMPRESION1,PUERTOIMPRESION2 Y PUERTOIMPRESION3
    If nroimpresion = 1 Or nroimpresion = 2 Or nroimpresion = 3 Then
        If Puerto = "%" Or Puerto = "" Then
            Exit Function

        End If

    End If
   
    ''11/07/2017 kenyo multicomandas

    If slpt = "COM1" Or slpt = "COM2" Or slpt = "COM3" Or slpt = "COM4" Or slpt = "COM5" Then
        'MsgBox slpt
        found = imprime_puerto_serial(slpt)
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
    For I = 1 To max
        Seek #1, I
        buf = input$(1, #1)
        Print #2, buf;
    Next I

    'MsgBox ""
    'End
    '---------------------------------------------
    Close #1
    Close #2
    cerrar_archivo
    Exit Function
cmd13081_err:

    MsgBox "Error en impresora " & I & " " & error$, 1, "Mensaje"

    Close #1
    Close #2
    Exit Function

End Function

Function corte_papel(apuerto As String, sw As Integer)

    Dim buf As String

    Dim I

    Dim found As Integer

    Select Case apuerto

        Case "LPT1", "LPT2", "LPT3", "LPT4", "LPT5"

            If sw = 0 Then   'star
                'buf = Chr$(27) + "d" + Chr$(0) 'star sp342
                I = FreeFile
                Open apuerto For Output As I
                'buf = Chr$(28) + Chr$(29)  'star sp200
                'buf = Chr$(7)   'star
                buf = Chr$(28) + Chr$(29)  'star sp200
                Print #I, buf;
                Close I

            End If

            If sw = 1 Then   'epson
                I = FreeFile
                Open apuerto For Output As I
                'buf = Chr$(27) + "i"  'epson
                'buf = Chr$(27) + "p" + Chr$(0) + Chr$(25) + Chr$(250) 'EPSON
                buf = Chr$(27) + "i"  'epson
                Print #I, buf;
                Close I

            End If

        Case "BAR"

            If sw = 0 Then
                buf = Chr$(28) + Chr$(29)  'star sp200

                'impresion_codbar buf
            End If

            If sw = 1 Then
                buf = Chr$(27) + "i"  'epson

                'impresion_codbar buf
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

Function abre_puerto(apuerto As String, _
                     sw As Integer, _
                     xcola As String) 'solo gaveta dinero

    Dim buf   As String

    Dim found As Integer

    Dim I

    On Error GoTo cmd8912_err

    If xcola = "S" Then
        found = gaveta_cola(apuerto)
        Exit Function

    End If

    Select Case apuerto

        Case "LPT1", "LPT2", "LPT3", "LPT4", "LPT5"

            If sw = 0 Then   'star
                'buf = Chr$(27) + "d" + Chr$(0) 'star sp342
                I = FreeFile
                Open apuerto For Output As I
                buf = Chr$(28) + Chr$(29)  'star sp200
                buf = Chr$(7)   'star
                Print #I, buf;
                Close I

            End If

            If sw = 1 Then   'epson
                I = FreeFile
                Open apuerto For Output As I
                buf = Chr$(27) + "i"  'epson
                buf = Chr$(27) + "p" + Chr$(0) + Chr$(25) + Chr$(250) 'EPSON
                Print #I, buf;
                Close I

            End If

        Case "BAR"

            If sw = 0 Then
                buf = Chr$(28) + Chr$(29)  'star sp200
                buf = Chr$(7)   'star

                'impresion_codbar buf
            End If

            If sw = 1 Then
                buf = Chr$(27) + "i"  'epson
                buf = Chr$(27) + "p" + Chr$(0) + Chr$(25) + Chr$(250) 'EPSON

                'impresion_codbar buf
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

Function valida_ruc(as_numero2 As String) As Integer

    Dim ic_num     As Integer

    Dim as_numero1 As String

    Dim xx         As Integer

    Dim as_numero  As String

    'valida_ruc = 1
    'Exit Function
    Dim lc_compara As Integer

    Dim lc_residuo As Integer

    Dim lc_valida  As Integer

    Dim found      As Integer

    'MsgBox as_numero2
    If Trim(menup.Label10) = "ARGENTINA" Then
        found = VerificarCUIT(as_numero2)
        valida_ruc = found
        Exit Function

    End If

    If Len(Trim(as_numero2)) = 11 Then
        valida_ruc = 1
        Exit Function

    End If

    '''' 17/07/2018 Factura de Exportación
    'If Len(as_numero2) <> 11 Then Exit Function
    If busca_OpcionExportacion = "N" Then
        If Len(as_numero2) <> 11 Then Exit Function

    End If

    '''' 17/07/2018 Factura de Exportación

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

Function imprime_delivery(mytablex As ADODB.Recordset)

    On Error GoTo cmd8912_err

    Dim buf      As String

    Dim mytablez As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim found    As Integer

    Dim dedonde  As String

    If Trim("" & mytablex.Fields("servicio")) <> "D" Then Exit Function
    'AQUI VEMOS DE DONDE VIENE SI ES MARKET O...
    dedonde = ""
    mytablez.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablez.RecordCount > 0 Then
        dedonde = "" & mytablez.Fields("touch")

    End If

    mytablez.Close

    If Len(Trim(dedonde)) = 0 Then  'si es de minimarket
        found = formateaa("", 1, 2, 0)
        buf = "----------------------------"
        found = formateaa(buf, 25, 0, 0)
        found = formateaa("", 1, 2, 0)
   
        found = formateaa("DATOS DELIVERY", 30, 2, 0)
        mytabley.Open "select * from clientes where codigo='" & "" & mytablex.Fields("codigo") & "'", cn, adOpenStatic, adLockOptimistic

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

            buf = imprime_clasifica_cliente(mytabley.Fields("codigo"))
            found = formateaa(buf, 30, 2, 0)

            '----------------------------
        End If

        mytabley.Close
        mytabley.Open "select * from factura where local='" & "" & mytablex.Fields("local") & "' and tipo='" & "" & mytablex.Fields("tipo") & "' and serie='" & "" & mytablex.Fields("serie") & "' and numero='" & "" & mytablex.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount > 0 Then
            buf = "Local  :" & mytabley.Fields("local")
            found = formateaa(buf, 15, 0, 0)
            found = formateaa("", 1, 2, 0)
            buf = "Tipo   :" & mytabley.Fields("tipo")
            found = formateaa(buf, 15, 0, 0)
            found = formateaa("", 1, 2, 0)
            buf = "Serie  :" & mytabley.Fields("Serie")
            found = formateaa(buf, 15, 0, 0)
            found = formateaa("", 1, 2, 0)
            buf = "Numero :" & mytabley.Fields("Numero")
            found = formateaa(buf, 15, 0, 0)
            found = formateaa("", 1, 2, 0)
            buf = "Cajero :" & mytabley.Fields("usuario")
            found = formateaa(buf, 15, 0, 0)
            found = formateaa("", 1, 2, 0)
            buf = "Caja   :" & mytabley.Fields("caja")
            found = formateaa(buf, 15, 0, 0)
            found = formateaa("", 1, 2, 0)
            buf = "Turno  :" & mytabley.Fields("turno")
            found = formateaa(buf, 15, 0, 0)
            found = formateaa("", 1, 2, 0)
            buf = "----------------------------"
            found = formateaa(buf, 25, 0, 0)
            found = formateaa("", 1, 2, 0)
            buf = "Driver :___________________"
            found = formateaa(buf, 25, 0, 0)
            found = formateaa("", 1, 2, 0)

        End If

        mytabley.Close
        'Forma Pago
        mytablez.Open "select * from fpagov where local='" & "" & mytablex.Fields("local") & "' and tipo='" & "" & mytablex.Fields("tipo") & "' and serie='" & "" & mytablex.Fields("serie") & "' and numero='" & "" & mytablex.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

        If mytablez.RecordCount > 0 Then
            Do

                If mytablez.EOF Then Exit Do
                buf = "Paga Con   :" & mytablez.Fields("Descripcio")
                found = formateaa(buf, 30, 2, 0)
                buf = "Monto      :" & mytablez.Fields("Recibe")
                found = formateaa(buf, 30, 2, 0)
                buf = "Vuelto     :" & mytablez.Fields("saldos")
                found = formateaa(buf, 30, 2, 0)
                mytablez.MoveNext
            Loop

        End If

        mytablez.Close
        Exit Function

    End If

    found = formateaa("DATOS DELIVERY", 30, 2, 0)
    mytabley.Open "select * from clientes where codigo='" & "" & mytablex.Fields("codigo") & "'", cn, adOpenStatic, adLockOptimistic

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

        buf = imprime_clasifica_cliente(mytabley.Fields("codigo"))
        found = formateaa(buf, 30, 2, 0)
          
        '----------------------------
    End If

    mytabley.Close
    'MsgBox "ABC"
    '----- datos del los producto
    mytabley.Open "select * from factura where local='" & "" & mytablex.Fields("local") & "' and tipo='" & "" & mytablex.Fields("tipo") & "' and serie='" & "" & mytablex.Fields("serie") & "' and numero='" & "" & mytablex.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        buf = "" & mytabley.Fields("tipo")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & imprime_tipodoc("" & mytabley.Fields("tipo"))
        found = formateaa(buf, 22, 2, 0)

    End If

    mytabley.Close
    'MsgBox "ABC"
    mytabley.Open "select * from detalle where local='" & "" & mytablex.Fields("local") & "' and tipo='" & "" & mytablex.Fields("tipo") & "' and serie='" & "" & mytablex.Fields("serie") & "' and numero='" & "" & mytablex.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        Do

            If mytabley.EOF Then Exit Do
            buf = "" & mytabley.Fields("cantidad")
            found = formateaa(buf, 6, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = Mid$("" & mytabley.Fields("descripcio"), 1, 22)
            found = formateaa(buf, 22, 2, 0)

            If Len(Mid$("" & mytabley.Fields("descripcio"), 23, 22)) > 0 Then
                buf = Mid$("" & mytabley.Fields("descripcio"), 23, 22)
                found = formateaa(buf, 22, 2, 0)

            End If

            If Len(Mid$("" & mytabley.Fields("descripcio"), 46, 22)) > 0 Then
                buf = Mid$("" & mytabley.Fields("descripcio"), 46, 22)
                found = formateaa(buf, 22, 2, 0)

            End If

            If Len(Mid$("" & mytabley.Fields("descripcio"), 69, 22)) > 0 Then
                buf = Mid$("" & mytabley.Fields("descripcio"), 69, 22)
                found = formateaa(buf, 22, 2, 0)

            End If

            If Len(Mid$("" & mytabley.Fields("descripcio"), 91, 22)) > 0 Then
                buf = Mid$("" & mytabley.Fields("descripcio"), 91, 22)
                found = formateaa(buf, 22, 2, 0)

            End If

            '----------------------
            If Len("" & mytabley.Fields("observa1")) > 0 Then
                buf = "*" & mytabley.Fields("observa1")
                found = formateaa(buf, 28, 2, 0)

            End If

            If Len("" & mytabley.Fields("observa2")) > 0 Then
                buf = "*" & mytabley.Fields("observa2")
                found = formateaa(buf, 28, 2, 0)

            End If

            If Len("" & mytabley.Fields("observa3")) > 0 Then
                buf = "*" & mytabley.Fields("observa3")
                found = formateaa(buf, 28, 2, 0)

            End If

            '----------------------
            mytabley.MoveNext
        Loop

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

    On Error GoTo errHandler

    FileCopy buf, buf1
    Exit Sub
errHandler:

    If Err = 55 Then    ' File already open.
        MsgBox "Tablas ya abiertas ,Salga y Vuelva a Ingresar. O limpie Temporales ", 24, "Aviso"
        End
    Else
        MsgBox "Por Favor Limpiar Temporales  Limpiar Temporales .", 48, "Aviso"
        End

    End If

    Resume Next

End Sub

Function inicializa_mscomm(D As Integer)

    On Error GoTo cmdini1_err

    menup.MSComm1.CommPort = D
    'MSComm1.InBufferSize = 1024
    menup.MSComm1.OutBufferSize = 512
    'menup.MSComm1.RThreshold = 15
    menup.MSComm1.SThreshold = 1
    ' MSComm1.InputLen = 15
    ' MSComm1.ParityReplace = "?"
    ' menup.MSComm1.RTSEnable = True
    ' menup.MSComm1.DTREnable = True
    ' menup.MSComm1.RTSEnable = True
    ' MSComm1.NullDiscard = False
    ' menup.MSComm1.Handshaking = comNone

    menup.MSComm1.Settings = "9600,n,8,1"

    If menup.MSComm1.PortOpen = True Then
        menup.MSComm1.PortOpen = False

    End If

    'menup.MSComm1.RThreshold = 1
    menup.MSComm1.PortOpen = True
    inicializa_mscomm = 1
    Exit Function
cmdini1_err:
    MsgBox "Aviso en inicializa mscomm " + error$, 48, "Aviso"
    Exit Function

End Function

Function escribe_mscomm(DATO As String) As Integer

    Dim I As Integer

    On Error GoTo cmdini2_err

    If menup.MSComm1.PortOpen = False Then Exit Function
    menup.MSComm1.Output = DATO
    escribe_mscomm = Len(DATO)

    Exit Function
cmdini2_err:
    Exit Function

End Function

Function cerrar_mscomm()

    On Error GoTo cmdini3_err

    If menup.MSComm1.PortOpen = False Then Exit Function
    menup.MSComm1.PortOpen = False
    cerrar_mscomm = 1
cmdini3_err:
    Exit Function

End Function

Sub cerrar_puertosmscomm()

    Dim I As Integer

    For I = 1 To 10
        cerrando_mscomm I
    Next I

End Sub

Sub cerrando_mscomm(D As Integer)

    On Error GoTo cmdini4_err

    menup.MSComm1.CommPort = D

    If menup.MSComm1.PortOpen = True Then
        menup.MSComm1.PortOpen = False

    End If

    Exit Sub
cmdini4_err:
    Exit Sub

End Sub

Sub Espera(Segundos As Single)

    Dim ComienzoSeg As Single

    Dim FinSeg      As Single

    ComienzoSeg = Timer
    FinSeg = ComienzoSeg + Segundos

    Do While FinSeg > Timer
        DoEvents

        If ComienzoSeg > Timer Then
            FinSeg = FinSeg - 24 * 60 * 60

        End If

    Loop

End Sub

Function kardexactualiza(xlocal1 As String, _
                         xproducto As String, _
                         xbodega As String, _
                         xfechai As String, _
                         xfechaf As String)

    Dim buf As String

    buf = " delete from talmacen " '  where producto like '" & xproducto & "'"
    'buf = buf & " and local='" & extra_loquesea(local1) & "'"
    'buf = buf & " and bodega='" & extra_loquesea(bodega) & "'"
    cn.Execute (buf)

    buf = "INSERT INTO tALMACEN (     local,producto, BODega,saldo) "

    buf = buf & " (SELECT     dsaldoini.local,dsaldoini.producto AS PROD,dsaldoini.bodega as bod, SUM(dsaldoini.cantidad*dsaldoini.factor) AS CANT"
    buf = buf & "                       From dsaldoini "
    buf = buf & "                       WHERE      "
    buf = buf & "   dsaldoini.fecha='" & Format(xfechai, "DD/MM/YYYY") & "'"
    buf = buf & " and dsaldoini.local='" & xlocal1 & "'"
    buf = buf & " and dsaldoini.producto like '" & xproducto & "'"

    buf = buf & " and dsaldoini.bodega='" & xbodega & "'"
    buf = buf & "                       GROUP BY dsaldoini.local,dsaldoini.producto,dsaldoini.bodega"

    buf = buf & "                       Union All "

    ''' 21/12/2017 Correccion de Recalculo de Saldo (En Nota de Credito)

    'buf = buf & "                       SELECT     detalle.local,detalle.producto AS PROD,detalle.bodega as bod, SUM(detalle.cantidad*detalle.factor) AS CANT"
    'buf = buf & "                       From detalle "
    'buf = buf & "  WHERE detalle.fecha>='" & Format(xfechai, "YYYYMMDD") & "'"
    'buf = buf & " and detalle.fecha<='" & Format(xfechaf, "YYYYMMDD") & "' "
    'buf = buf & " and detalle.local='" & xlocal1 & "'"
    'buf = buf & " and detalle.producto like '" & xproducto & "'"
    'buf = buf & " and detalle.estado='2' AND detalle.acu1='' "
    'buf = buf & " and detalle.bodega='" & xbodega & "'"
    'buf = buf & "                        AND (detalle.acu = 'J' OR"
    'buf = buf & "                                             detalle.acu = 'K' OR"
    'buf = buf & "                                             detalle.acu = 'L' OR"
    'buf = buf & "                                             detalle.acu = 'M' OR"
    'buf = buf & "                                             detalle.acu = 'P' or detalle.acu='S' or detalle.acu='E')"
    'buf = buf & "                       GROUP BY detalle.local,detalle.producto,detalle.bodega "
    'buf = buf & "                       Union All "

    'buf = buf & "                       SELECT     detalle.local,detalle.producto AS PROD,detalle.bodega as bod, - SUM(detalle.cantidad*detalle.factor) AS CANT"
    'buf = buf & "                       From detalle "
    'buf = buf & "  WHERE detalle.fecha>='" & Format(xfechai, "YYYYMMDD") & "'"
    'buf = buf & " and detalle.fecha<='" & Format(xfechaf, "YYYYMMDD") & "' "
    'buf = buf & " and detalle.local='" & xlocal1 & "'"
    'buf = buf & " and detalle.producto like '" & xproducto & "'"
    'buf = buf & " and detalle.estado='2' AND detalle.acu1='' "
    'buf = buf & " and detalle.bodega='" & xbodega & "'"
    'buf = buf & "                        AND (detalle.acu = 'A' OR"
    'buf = buf & "                                             detalle.acu = 'B' OR"
    'buf = buf & "                                             detalle.acu = 'C' OR"
    'buf = buf & "                                             detalle.acu = 'D' OR"
    'buf = buf & "                                             detalle.acu = 'G' or detalle.acu='T' or detalle.acu='N')"
    'buf = buf & "                       GROUP BY detalle.local,detalle.producto,detalle.bodega) "
    '

    buf = buf & "                       SELECT     detalle.local,detalle.producto AS PROD,detalle.bodega as bod, SUM(detalle.cantidad*detalle.factor) AS CANT"
    buf = buf & "                       From detalle "
    buf = buf & "  WHERE detalle.fecha>='" & Format(xfechai, "YYYYMMDD") & "'"
    buf = buf & " and detalle.fecha<='" & Format(xfechaf, "YYYYMMDD") & "' "
    buf = buf & " and detalle.local='" & xlocal1 & "'"
    buf = buf & " and detalle.producto like '" & xproducto & "'"
    buf = buf & " and detalle.estado='2' AND detalle.acu1='' "
    buf = buf & " and detalle.bodega='" & xbodega & "'"
    buf = buf & "                        AND (detalle.acu = 'J' OR"
    buf = buf & "                                             detalle.acu = 'K' OR"
    buf = buf & "                                             detalle.acu = 'L' OR"
    buf = buf & "                                             detalle.acu = 'M' OR"

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
    buf = buf & "                                             detalle.acu = 'E' OR"
    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    buf = buf & "                                             detalle.acu = 'P' or detalle.acu='S')"
    buf = buf & "                       GROUP BY detalle.local,detalle.producto,detalle.bodega "
    buf = buf & "                       Union All "

    buf = buf & "                       SELECT     detalle.local,detalle.producto AS PROD,detalle.bodega as bod, - SUM(detalle.cantidad*detalle.factor) AS CANT"
    buf = buf & "                       From detalle "
    buf = buf & "  WHERE detalle.fecha>='" & Format(xfechai, "YYYYMMDD") & "'"
    buf = buf & " and detalle.fecha<='" & Format(xfechaf, "YYYYMMDD") & "' "
    buf = buf & " and detalle.local='" & xlocal1 & "'"
    buf = buf & " and detalle.producto like '" & xproducto & "'"
    buf = buf & " and detalle.estado='2' AND detalle.acu1='' "
    buf = buf & " and detalle.bodega='" & xbodega & "'"
    buf = buf & "                        AND (detalle.acu = 'A' OR"
    buf = buf & "                                             detalle.acu = 'B' OR"
    buf = buf & "                                             detalle.acu = 'C' OR"
    buf = buf & "                                             detalle.acu = 'D' OR"

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
    'buf = buf & "                                             detalle.acu = 'G' or detalle.acu='E' or detalle.acu='T' or detalle.acu='N')"
    buf = buf & "                                             detalle.acu = 'G' or detalle.acu='F' or detalle.acu='T' or detalle.acu='N')"

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    buf = buf & "                       GROUP BY detalle.local,detalle.producto,detalle.bodega) "

    ''' 21/12/2017 Correccion de Recalculo de Saldo (En Nota de Credito)

    'MsgBox buf
    cn.Execute (buf)

    'buf = " delete from almacen  where producto like '" & xproducto & "'"
    'buf = buf & " and local='" & xlocal1 & "'"
    'buf = buf & " and bodega='" & xbodega & "'"
    'MsgBox buf
    buf = " delete from almacen  where  " 'producto like '" & xproducto & "'"
    buf = buf & " local='" & xlocal1 & "'"
    buf = buf & " and producto like '" & xproducto & "'"
    buf = buf & " and bodega='" & xbodega & "'"
    cn.Execute (buf)

    buf = "insert into almacen (local,producto,bodega,saldo) (select talmacen.local,talmacen.producto,talmacen.bodega,sum(talmacen.saldo) as saldo "
    buf = buf & " from talmacen   "
    buf = buf & " group by local,producto,bodega)"

    'buf = "insert into almacen select * from talmacen to,almacen td where "
    'buf = " to.local<>td.local"
    'buf = " and to.producto<>td.producto"
    'buf = "   and to.bodega<>td.bodega"
    cn.Execute (buf)

    'buf = "INSERT INTO almacen "
    'buf = buf & " (producto,local,bodega,saldo,entrada,salida,saldoinicial)"
    'buf = buf & "SELECT     PRODUCTO,'" & extra_loquesea(local1) & "','" & extra_loquesea(bodega) & "',0,0,0,0"
    'buf = buf & " From producto  "
    'buf = buf & " WHERE   producto like '" & producto & "'"
    'cn.Execute (buf)
    'Exit Function

    'buf = "INSERT INTO almacen "
    'buf = buf & " (producto,local,bodega,saldo)"
    'buf = buf & "SELECT sum(    PRODUCTO,'" & extra_loquesea(local1) & "','" & extra_loquesea(bodega) & "',0,0,0,0"
    'buf = buf & " From producto  "
    'buf = buf & " WHERE   producto like '" & producto & "'"
    'cn.Execute (buf)

    'buf = "update almacen set entrada=("
    'buf = buf & "SELECT   sum(CANTIDAD*factor)"
    'buf = buf & " From detalle  "
    'buf = buf & " WHERE detalle.producto=almacen.producto and (DETALLE.ACU = 'J' OR  DETALLE.ACU = 'K' OR DETALLE.ACU = 'L' OR  DETALLE.ACU = 'M' OR  DETALLE.ACU = 'S') and DETALLE.estado='2' "
    'buf = buf & " and detalle.bodega=almacen.bodega and almacen.local=detalle.local "
    'buf = buf & "  and detalle.fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    'buf = buf & " and detalle.fecha<='" & Format(fechaf, "YYYYMMDD") & "') "
    'cn.Execute (buf)

    'buf = "update almacen set salida=("
    'buf = buf & "SELECT   sum(CANTIDAD*factor)"
    'buf = buf & " From detalle  "
    'buf = buf & " WHERE detalle.producto=almacen.producto and (DETALLE.ACU = 'A' OR  DETALLE.ACU = 'B' OR DETALLE.ACU = 'C' OR  DETALLE.ACU = 'D' OR  DETALLE.ACU = 'G') and DETALLE.estado='2' "
    'buf = buf & " and detalle.bodega=almacen.bodega and almacen.local=detalle.local "
    'buf = buf & "  and detalle.fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    'buf = buf & " and detalle.fecha<='" & Format(fechaf, "YYYYMMDD") & "') "
    'cn.Execute (buf)

    'buf = "update almacen set saldoinicial=("
    'buf = buf & "SELECT   sum(CANTIDAD*factor)"
    'buf = buf & " From saldoini  "
    'buf = buf & " WHERE saldoini.producto=almacen.producto  "
    'buf = buf & " and saldoini.bodega=almacen.bodega and almacen.local=saldoini.local "
    'buf = buf & "  and saldoini.fecha='" & Format(fechai, "YYYYMMDD") & "')"
    'cn.Execute (buf)

    'buf = "update almacen set saldo=entrada-salida+saldoinicial "
    'cn.Execute (buf)
    'MsgBox "Proceso terminado ", 48, "Aviso"
    kardexactualiza = 1

End Function

Sub borra_tabla(buf As String)

    On Error GoTo cmd9067_err

    cn.Execute (buf)
    Exit Sub
cmd9067_err:
    MsgBox "Aviso en borra Tabla " + error$, 48, "Aviso"
    Exit Sub

End Sub

Function leer_visorcaja(buf1 As String, buf2 As String)

    Dim found    As Integer

    Dim bdvisor1 As bdvisor

    On Error GoTo cmd7824_err

    Dim buf As String

    Open globalpath & "\visor.txt" For Random As #4 Len = Len(bdvisor1)
    Get #4, 1, bdvisor1
    found = envio_visor(bdvisor1.ppuerto, bdvisor1.vvelocidad, buf1, buf2)
    leer_visorcaja = 1
    Close #4
    Exit Function
cmd7824_err:
    MsgBox "Aviso en Leer visor " + error$, 48, "Aviso"
    Exit Function

End Function

Function envio_visor(kpuerto As String, _
                     kvelocidad As String, _
                     kmensaje1 As String, _
                     kmensaje2 As String)

    Dim I As Integer

    On Error GoTo cmd9090_err

    menup.visorcl.CommPort = kpuerto
    menup.visorcl.Settings = kvelocidad

    If menup.visorcl.PortOpen = True Then
        menup.visorcl.PortOpen = False

    End If

    menup.visorcl.PortOpen = True
    menup.visorcl.Output = Chr$(27) & Chr$(81) & Chr$(65)
    menup.visorcl.Output = Mid$(Trim(kmensaje1), 1, 16) & Chr$(13)
    menup.visorcl.Output = Chr$(27) & Chr$(81) & Chr$(66)
    menup.visorcl.Output = Mid$(Trim(kmensaje2), 1, 16) & Chr$(13)
    menup.visorcl.PortOpen = False
    Exit Function
cmd9090_err:

    If menup.visorcl.PortOpen = True Then
        menup.visorcl.PortOpen = False

    End If

    'MsgBox "Aviso en envio Visor ", 48, "Aviso"
    Exit Function

End Function

Function buscar_list(lst As ListBox, zString As String)

    Dim I As Integer

    On Error Resume Next

    buscar_list = 1

    For I = 0 To lst.ListCount

        If Mid$(lst.List(I), 1, Len(zString)) = zString Then
            'LstIsIn = True
            lst.ListIndex = I
            GoTo grr

        End If

    Next I

    'ListIsIn = 0
grr:

End Function

Private Sub espera_segundo(ByVal nSec As Integer)

    'Esperar un número de segundos
    Dim t1 As Date, t2 As Date

    t1 = Second(Now)
    t2 = t1 + nSec
    Do
        DoEvents
    Loop While t2 > Second(Now)

End Sub

Sub SaveBitmap(mytablex As ADODB.Recordset, SourceFile As String)

    On Error GoTo cmd9090_err

    Dim Arr()         As Byte

    Dim Pointer       As Long

    Dim SizeOfThefile As Long

    Dim buf           As String

    'buf = globaldir & "\grafico\" & SourceFile
    Pointer = lOpen(SourceFile, OF_READ)
    'size of the file
    SizeOfThefile = GetFileSize(Pointer, lpFSHigh)
    lclose Pointer

    'Resize the array, then fill it with
    'the entire contents of the field
    
    '' 02/12/2017 Opcion a poder borrar imagen del modulo de productos
    If tproduct.fotonombre = "" Then
        Close #1
        mytablex("imagen").Value = ""

    End If

    '' 02/12/2017 Opcion a poder borrar imagen del modulo de productos
    
    ReDim Arr(SizeOfThefile)

    Open SourceFile For Binary Access Read As #1
    Get #1, , Arr
    Close #1
    mytablex("imagen").Value = Arr
    Exit Sub
cmd9090_err:
    'MsgBox "Avis en savebitmap ", 48, "Aviso"
    Exit Sub
    
End Sub

Sub viewBMP(mytablex As ADODB.Recordset, fotonombre As String)

    Dim lenf    As Long

    Dim Offset  As Long

    Dim chunk() As Byte

    Dim buf     As String

    On Error GoTo cmd9912_err

    Offset = 0
    borrar_archivo fotonombre
    buf = fotonombre
    lenf = mytablex.Fields("imagen").ActualSize

    If lenf > 0 Then
        Open buf For Binary As 1

        Do Until Offset >= lenf
            chunk() = mytablex.Fields("imagen").GetChunk(100)
            Put #1, , chunk()
            Offset = Offset + 100
        Loop
        Close #1

    End If

    Exit Sub
cmd9912_err:
    MsgBox "Aviso en ViewBMP " + error$, 48, "Aviso"
    Exit Sub

End Sub

Function Verificar_ruc(ByVal xnum As String) As Boolean

    Dim li_suma, li_residuo, li_diferencia, li_compara As Integer

    li_suma = (CInt(Mid(xnum, 1, 1)) * 5) + (CInt(Mid(xnum, 2, 1)) * 4) + (CInt(Mid(xnum, 3, 1)) * 3) + (CInt(Mid(xnum, 4, 1)) * 2) + (CInt(Mid(xnum, 5, 1)) * 7) + (CInt(Mid(xnum, 6, 1)) * 6) + (CInt(Mid(xnum, 7, 1)) * 5) + (CInt(Mid(xnum, 8, 1)) * 4) + (CInt(Mid(xnum, 9, 1)) * 3) + (CInt(Mid(xnum, 10, 1)) * 2)
    li_compara = CInt(Mid(xnum, 11, 1))
    li_residuo = li_suma Mod 11
    li_diferencia = Int(11 - li_residuo)

    If li_diferencia > 9 Then li_diferencia = li_diferencia - 10
    If li_diferencia <> li_compara Then
        Verificar_ruc = False
    Else
        Verificar_ruc = True

    End If

End Function

Function valida_numero_ruc(buf As String) As Boolean

    On Error GoTo cmd9090_err

    If Val(Mid(Trim(buf), 2, 9)) = 0 Or Trim(buf) = "23333333333" Then
        Exit Function

    End If

    If Verificar_ruc(buf) = False Then
        Exit Function

    End If

    valida_numero_ruc = True
    Exit Function
cmd9090_err:
    valida_numero_ruc = False
    Exit Function

End Function

Function imprime_tipodoc(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "Select * from tipo where tipo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        imprime_tipodoc = Trim("" & mytablex.Fields("descripcio"))

    End If

    mytablex.Close

End Function

Function imprime_clasifica_cliente(buf As String)

    Dim mytablex As New ADODB.Recordset

    'MsgBox buf
    mytablex.Open "Select * from clientes where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        imprime_clasifica_cliente = Trim("" & mytablex.Fields("clasifica"))

    End If

    mytablex.Close

End Function

Function kardexactualizasi(xlocal1 As String, _
                           xproducto As String, _
                           xbodega As String, _
                           xfechai As String, _
                           xfechaf As String)

    Dim buf As String

    buf = " delete from talmacen " '  where producto like '" & xproducto & "'"
    cn.Execute (buf)
    buf = " delete from almacen0 " '  where producto like '" & xproducto & "'"
    cn.Execute (buf)

    buf = "INSERT INTO tALMACEN (     local,producto, BODega,saldo) "
    buf = buf & " (SELECT     saldoini.local,saldoini.producto AS PROD,saldoini.bodega as bod, SUM(saldoini.cantidad1*saldoini.factor) AS CANT"
    buf = buf & "                       From saldoini "
    buf = buf & "                       WHERE      "
    buf = buf & "   saldoini.fecha='" & Format(xfechai, "DD/MM/YYYY") & "'"
    buf = buf & " and saldoini.local='" & xlocal1 & "'"
    buf = buf & " and saldoini.producto like '" & xproducto & "'"

    buf = buf & " and saldoini.bodega='" & xbodega & "'"
    buf = buf & "                       GROUP BY saldoini.local,saldoini.producto,saldoini.bodega"

    buf = buf & "                       Union All "
    buf = buf & "                       SELECT     detalle.local,detalle.producto AS PROD,detalle.bodega as bod, SUM(detalle.cantidad*detalle.factor) AS CANT"
    buf = buf & "                       From detalle "
    buf = buf & "  WHERE detalle.fecha>='" & Format(xfechai, "YYYYMMDD") & "'"
    buf = buf & " and detalle.fecha<='" & Format(xfechaf, "YYYYMMDD") & "' "
    buf = buf & " and detalle.local='" & xlocal1 & "'"
    buf = buf & " and detalle.producto like '" & xproducto & "'"
    buf = buf & " and detalle.estado='2' AND detalle.acu1='' "
    buf = buf & " and detalle.bodega='" & xbodega & "'"
    buf = buf & "                        AND (detalle.acu = 'J' OR"
    buf = buf & "                                             detalle.acu = 'K' OR"
    buf = buf & "                                             detalle.acu = 'L' OR"
    buf = buf & "                                             detalle.acu = 'M' OR"
    buf = buf & "                                             detalle.acu = 'P' or detalle.acu='S' or detalle.acu='E')"
    buf = buf & "                       GROUP BY detalle.local,detalle.producto,detalle.bodega "
    buf = buf & "                       Union All "
    buf = buf & "                       SELECT     detalle.local,detalle.producto AS PROD,detalle.bodega as bod, - SUM(detalle.cantidad*detalle.factor) AS CANT"
    buf = buf & "                       From detalle "
    buf = buf & "  WHERE detalle.fecha>='" & Format(xfechai, "YYYYMMDD") & "'"
    buf = buf & " and detalle.fecha<='" & Format(xfechaf, "YYYYMMDD") & "' "
    buf = buf & " and detalle.local='" & xlocal1 & "'"
    buf = buf & " and detalle.producto like '" & xproducto & "'"
    buf = buf & " and detalle.estado='2' AND detalle.acu1='' "
    buf = buf & " and detalle.bodega='" & xbodega & "'"
    buf = buf & "                        AND (detalle.acu = 'A' OR"
    buf = buf & "                                             detalle.acu = 'B' OR"
    buf = buf & "                                             detalle.acu = 'C' OR"
    buf = buf & "                                             detalle.acu = 'D' OR"
    buf = buf & "                                             detalle.acu = 'G' or detalle.acu='T' or detalle.acu='N')"
    buf = buf & "                       GROUP BY detalle.local,detalle.producto,detalle.bodega) "
    'MsgBox buf
    cn.Execute (buf)

    'buf = " delete from almacen  where producto like '" & xproducto & "'"
    'buf = buf & " and local='" & xlocal1 & "'"
    'buf = buf & " and bodega='" & xbodega & "'"
    'MsgBox buf
    buf = " delete from almacen0  where  " 'producto like '" & xproducto & "'"
    buf = buf & " local='" & xlocal1 & "'"
    buf = buf & " and producto like '" & xproducto & "'"
    buf = buf & " and bodega='" & xbodega & "'"
    cn.Execute (buf)

    buf = "insert into almacen0 (local,producto,bodega,saldo) (select talmacen.local,talmacen.producto,talmacen.bodega,sum(talmacen.saldo) as saldo "
    buf = buf & " from talmacen   "
    buf = buf & " group by local,producto,bodega)"

    'buf = "insert into almacen select * from talmacen to,almacen td where "
    'buf = " to.local<>td.local"
    'buf = " and to.producto<>td.producto"
    'buf = "   and to.bodega<>td.bodega"
    cn.Execute (buf)

    'buf = "INSERT INTO almacen "
    'buf = buf & " (producto,local,bodega,saldo,entrada,salida,saldoinicial)"
    'buf = buf & "SELECT     PRODUCTO,'" & extra_loquesea(local1) & "','" & extra_loquesea(bodega) & "',0,0,0,0"
    'buf = buf & " From producto  "
    'buf = buf & " WHERE   producto like '" & producto & "'"
    'cn.Execute (buf)
    'Exit Function

    'buf = "INSERT INTO almacen "
    'buf = buf & " (producto,local,bodega,saldo)"
    'buf = buf & "SELECT sum(    PRODUCTO,'" & extra_loquesea(local1) & "','" & extra_loquesea(bodega) & "',0,0,0,0"
    'buf = buf & " From producto  "
    'buf = buf & " WHERE   producto like '" & producto & "'"
    'cn.Execute (buf)

    'buf = "update almacen set entrada=("
    'buf = buf & "SELECT   sum(CANTIDAD*factor)"
    'buf = buf & " From detalle  "
    'buf = buf & " WHERE detalle.producto=almacen.producto and (DETALLE.ACU = 'J' OR  DETALLE.ACU = 'K' OR DETALLE.ACU = 'L' OR  DETALLE.ACU = 'M' OR  DETALLE.ACU = 'S') and DETALLE.estado='2' "
    'buf = buf & " and detalle.bodega=almacen.bodega and almacen.local=detalle.local "
    'buf = buf & "  and detalle.fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    'buf = buf & " and detalle.fecha<='" & Format(fechaf, "YYYYMMDD") & "') "
    'cn.Execute (buf)

    'buf = "update almacen set salida=("
    'buf = buf & "SELECT   sum(CANTIDAD*factor)"
    'buf = buf & " From detalle  "
    'buf = buf & " WHERE detalle.producto=almacen.producto and (DETALLE.ACU = 'A' OR  DETALLE.ACU = 'B' OR DETALLE.ACU = 'C' OR  DETALLE.ACU = 'D' OR  DETALLE.ACU = 'G') and DETALLE.estado='2' "
    'buf = buf & " and detalle.bodega=almacen.bodega and almacen.local=detalle.local "
    'buf = buf & "  and detalle.fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    'buf = buf & " and detalle.fecha<='" & Format(fechaf, "YYYYMMDD") & "') "
    'cn.Execute (buf)

    'buf = "update almacen set saldoinicial=("
    'buf = buf & "SELECT   sum(CANTIDAD*factor)"
    'buf = buf & " From saldoini  "
    'buf = buf & " WHERE saldoini.producto=almacen.producto  "
    'buf = buf & " and saldoini.bodega=almacen.bodega and almacen.local=saldoini.local "
    'buf = buf & "  and saldoini.fecha='" & Format(fechai, "YYYYMMDD") & "')"
    'cn.Execute (buf)

    'buf = "update almacen set saldo=entrada-salida+saldoinicial "
    'cn.Execute (buf)
    'MsgBox "Proceso terminado ", 48, "Aviso"
    kardexactualizasi = 1

End Function

Function redondeo1(buf3 As String, nrodecimal1 As String) As String

    Dim buf0 As String

    Dim buf1 As String

    Dim buf2 As String

    Dim sdx  As Double

    Dim buf  As String

    Dim c    As Double

    Dim D    As Double

    buf = buf3
    c = Val(buf3)
    D = Round(c, 2)
    redondeo1 = "" & D
    Exit Function
    buf = Format(Val(buf), nrodecimal1)

    buf0 = Mid$(buf, 1, Len(buf) - 3)
    buf1 = Mid$(buf, Len(buf) - 1, 2)
    buf2 = ""

    If Val(buf1) <= 0 Then
        redondeo1 = buf3

    End If

    'MsgBox buf1

    '00
    '02
    '03
    '04
    '05
    '06
    '07
    '08
    '09

    If Val(Mid$(buf1, 1, 1)) = 9 And Val(Mid$(buf1, 2, 1)) > 5 And Val(Mid$(buf1, 2, 1)) <= 9 Then
        sdx = Val(Mid$(buf0, 1, 1)) + 1
        buf2 = Format(sdx, "0")
        buf = buf2
        buf = Format(Val(buf), nrodecimal1)
        redondeo1 = buf
        Exit Function

    End If

    If Val(Mid$(buf1, 2, 1)) >= 0 And Val(Mid$(buf1, 2, 1)) <= 5 Then
        buf2 = Mid$(buf1, 1, 1) & "5"
        'buf2 = Mid$(buf1, 1, 1) & "0"
        buf = buf0 + "." + buf2
        buf = Format(Val(buf), nrodecimal1)
        redondeo1 = buf
        Exit Function

    End If

    If Val(Mid$(buf1, 2, 1)) > 5 And Val(Mid$(buf1, 2, 1)) <= 9 Then
        sdx = Val(Mid$(buf1, 1, 1)) + 1
        buf2 = Format(sdx, "0")
        buf = buf0 & "." & buf2
        buf = Format(Val(buf), nrodecimal1)
        redondeo1 = buf
        Exit Function

    End If

    redondeo1 = buf3

End Function

Public Function Redondear1(dblnToR As Double, Optional intCntDec As Integer) As Double

    Dim dblPot As Double

    Dim dblF   As Double
    
    If dblnToR < 0 Then dblF = -0.5 Else: dblF = 0.5
    dblPot = 10 ^ intCntDec
    Redondear1 = Fix(dblnToR * dblPot * (1 + 1E-16) + dblF) / dblPot

End Function

Function OTROPOS(ByVal xnum As String, xdir As String) As String

    Dim xDat    As String

    Dim xRazSoc As String

    On Error Resume Next

    Dim xWml As New XMLHTTP

    xWml.Open "POST", "http://www.sunat.gob.pe/w/wapS01Alias?ruc=" & xnum, False
    xWml.send

    If xWml.Status = 200 Then
        'Limpiar
        xDat = xWml.responseText

        If Len(xDat) <= 635 Then
            'Habilitar False
            MsgBox "El numero Ruc ingresado no existe en la Base de datos de la SUNAT", 48, "Aviso"
            Set xWml = Nothing
            'txtruc.SetFocus
            Exit Function

        End If

        Dim xTabla() As String

        xdir = ""
        xDat = Replace(xDat, "     ", " ")
        xDat = Replace(xDat, "    ", " ")
        xDat = Replace(xDat, "   ", " ")
        xDat = Replace(xDat, "  ", " ")
        xDat = Replace(xDat, "( ", "(")
        xDat = Replace(xDat, " )", ")")
       
        xTabla = Split(xDat, "<small>")
      
        xTabla(1) = Replace(xTabla(1), "<b>N&#xFA;mero Ruc. </b> " & xnum & " - ", "")
        xTabla(1) = Replace(xTabla(1), " <br/></small>", "")
       
        xTabla(4) = Replace(xTabla(4), "<b>Estado.</b>", "")
        xTabla(4) = Replace(xTabla(4), "</small><br/>", "")
       
        xTabla(7) = Replace(xTabla(7), "<b>Direcci&#xF3;n.</b><br/>", "")
        xTabla(7) = Replace(xTabla(7), "</small><br/>", "")
       
        xTabla(8) = Replace(xTabla(8), "Situaci&#xF3;n.<b> ", "")
        xTabla(8) = Replace(xTabla(8), "</b></small><br/>", "")
       
        xRazSoc = CStr(xTabla(1))
        'xEst = CStr(xTabla(4))
        xdir = CStr(xTabla(7))
        'xCon = CStr(xTabla(8))
       
        xRazSoc = Replace(xRazSoc, "&#209;", "Ñ")
        xRazSoc = Replace(xRazSoc, "&#xD1;", "Ñ")
        xRazSoc = Replace(xRazSoc, "&#193;", "Á")
        xRazSoc = Replace(xRazSoc, "&#201;", "É")
        xRazSoc = Replace(xRazSoc, "&#205;", "Í")
        xRazSoc = Replace(xRazSoc, "&#211;", "Ó")
        xRazSoc = Replace(xRazSoc, "&#218;", "Ú")
        xRazSoc = Replace(xRazSoc, "&#xC1;", "Á")
        xRazSoc = Replace(xRazSoc, "&#xC9;", "É")
        xRazSoc = Replace(xRazSoc, "&#xCD;", "Í")
        xRazSoc = Replace(xRazSoc, "&#xD3;", "Ó")
        xRazSoc = Replace(xRazSoc, "&#xDA;", "Ú")
       
        xRazSoc = Mid(xRazSoc, 1, Len(xRazSoc) - 3)
        'direccion
        xdir = Replace(xdir, "&#209;", "Ñ")
        xdir = Replace(xdir, "&#xD1;", "Ñ")
        xdir = Replace(xdir, "&#193;", "Á")
        xdir = Replace(xdir, "&#201;", "É")
        xdir = Replace(xdir, "&#205;", "Í")
        xdir = Replace(xdir, "&#211;", "Ó")
        xdir = Replace(xdir, "&#218;", "Ú")
        xdir = Replace(xdir, "&#xC1;", "Á")
        xdir = Replace(xdir, "&#xC9;", "É")
        xdir = Replace(xdir, "&#xCD;", "Í")
        xdir = Replace(xdir, "&#xD3;", "Ó")
        xdir = Replace(xdir, "&#xDA;", "Ú")
        xdir = Trim(Mid$(Mid(xdir, 1, Len(xdir) - 3), 1, 200))
        'MsgBox xdir
        'MsgBox xRazSoc
        OTROPOS = Trim(xRazSoc)
        
    Else
        MsgBox "No responde el servicio de la SUNAT", 48, "Aviso"

    End If

    Set xWml = Nothing

End Function

Function redondeo2(buf3 As String, nrodecimal1 As String) As String

    Dim buf0 As String

    Dim buf1 As String

    Dim buf2 As String

    Dim sdx  As Double

    Dim buf  As String

    Dim c    As Double

    Dim D    As Double

    On Error GoTo cmd5555_err

    buf = buf3
    'c = Val(buf3)
    'd = Round(c, 2)
    'redondeo1 = "" & d
    'Exit Function
    buf = Format(Val(buf), nrodecimal1)
    buf0 = Mid$(buf, 1, Len(buf) - 3)
    ''MsgBox "" & buf & ".. " & buf0
    buf1 = Mid$(buf, Len(buf) - 1, 2)
    buf2 = ""

    If Val(buf1) <= 0 Then
        redondeo2 = buf3
        'MsgBox buf3
        Exit Function

    End If

    'MsgBox buf0

    '00
    '02
    '03
    '04
    '05
    '06
    '07
    '08
    '09
    'MsgBox "xx " & buf1

    'MsgBox buf0

    ''18/07/2017 kenyo Redondeo decimales al vender granel
    'If Val(Mid$(buf1, 1, 1)) = 9 And Val(Mid$(buf1, 2, 1)) > 5 And Val(Mid$(buf1, 2, 1)) <= 9 Then
    If Val(Mid$(buf1, 1, 1)) = 9 And Val(Mid$(buf1, 2, 1)) >= 5 And Val(Mid$(buf1, 2, 1)) <= 9 Then
 
        ''18/07/2017 kenyo Redondeo decimales al vender granel
      
        sdx = Val(Mid$(buf0, 1, 1)) + 1
        sdx = Val(buf0) + 1
        buf2 = Format(sdx, "0")
        buf = buf2
        buf = Format(Val(buf2), nrodecimal1)
        redondeo2 = buf
        'MsgBox "hu " & buf0 ' & " " & buf
        Exit Function

    End If

    ''18/07/2017 kenyo Redondeo decimales al vender granel

    'If Val(Mid$(buf1, 2, 1)) >= 0 And Val(Mid$(buf1, 2, 1)) <= 5 Then
    If Val(Mid$(buf1, 2, 1)) >= 0 And Val(Mid$(buf1, 2, 1)) < 5 Then
 
        ''18/07/2017 kenyo Redondeo decimales al vender granel
   
        'buf2 = Mid$(buf1, 1, 1) & "5"
        buf2 = Mid$(buf1, 1, 1) & "0"
        buf = buf0 + "." + buf2
        buf = Format(Val(buf), nrodecimal1)
        redondeo2 = buf
        Exit Function

    End If

    ''18/07/2017 kenyo Redondeo decimales al vender granel
    'If Val(Mid$(buf1, 2, 1)) > 5 And Val(Mid$(buf1, 2, 1)) <= 9 Then
    If Val(Mid$(buf1, 2, 1)) >= 5 And Val(Mid$(buf1, 2, 1)) <= 9 Then

        ''18/07/2017 kenyo Redondeo decimales al vender granel
   
        sdx = Val(Mid$(buf1, 1, 1)) + 1
        buf2 = Format(sdx, "0")
        buf = buf0 & "." & buf2
        buf = Format(Val(buf), nrodecimal1)
        redondeo2 = buf
   
        Exit Function

    End If

    ''18/07/2017 kenyo Redondeo decimales al vender granel

    'MsgBox "xx"
    redondeo2 = buf3
    Exit Function
cmd5555_err:
    MsgBox "Aviso en redondeo2 " + error, 48, "Aviso"
    Exit Function

End Function

Function selecciona_percepcion(buf As String, buf2 As String) As Double

    Dim mytablex As New ADODB.Recordset

    Dim buf1     As String

    On Error GoTo cmd90123_err

    If Len(Trim("" & buf)) = 11 Then
        If valida_numero_ruc("" & buf) = True Then  'si es ruc verdarr
            buf1 = "SELECT * from clasesunat where clasesunat='" & extra_loquesea(buf2) & "'"
            mytablex.Open buf1, cn, adOpenDynamic, adLockOptimistic

            If mytablex.RecordCount > 0 Then
                selecciona_percepcion = Val("" & mytablex.Fields("percepcion"))

            End If

            mytablex.Close

        End If

    End If

    Exit Function
cmd90123_err:
    MsgBox "Aviso en selecciona percepcion " + error$, 48, "Aviso"
    Exit Function

End Function

Public Function SendMail(sTo As String, _
                         sSubject As String, _
                         sFrom As String, _
                         sBody As String, _
                         sSmtpServer As String, _
                         iSmtpPort As Integer, _
                         sSmtpUser As String, _
                         sSmtpPword As String, _
                         sFilePath As String, _
                         bSmtpSSL As Boolean, _
                         txtselecciona As String, _
                         txthtml1 As String) As String
    
    Dim txthtml As String

    Dim fso, ts

    On Error GoTo SendMail_Error:

    Dim lobj_cdomsg As CDO.Message

    Set lobj_cdomsg = New CDO.Message

    'MsgBox txtselecciona
    If txtselecciona = "H" Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set ts = fso.OpenTextFile(txthtml1)
        txthtml = ts.ReadAll

    End If
    
    lobj_cdomsg.Configuration.Fields(cdoSMTPServer) = sSmtpServer
    lobj_cdomsg.Configuration.Fields(cdoSMTPServerPort) = iSmtpPort
    lobj_cdomsg.Configuration.Fields(cdoSMTPUseSSL) = bSmtpSSL
    lobj_cdomsg.Configuration.Fields(cdoSMTPAuthenticate) = cdoBasic
    lobj_cdomsg.Configuration.Fields(cdoSendUserName) = sSmtpUser
    lobj_cdomsg.Configuration.Fields(cdoSendPassword) = sSmtpPword
    lobj_cdomsg.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
    lobj_cdomsg.Configuration.Fields(cdoSendUsingMethod) = cdoSendUsingPort
    lobj_cdomsg.Configuration.Fields.Update
    lobj_cdomsg.To = sTo
    lobj_cdomsg.From = sFrom
    lobj_cdomsg.Subject = sSubject

    If txtselecciona = "T" Then
        lobj_cdomsg.TextBody = sBody

    End If

    If txtselecciona = "H" Then
        lobj_cdomsg.HTMLBody = txthtml

    End If

    If Trim$(sFilePath) <> vbNullString Then
        lobj_cdomsg.AddAttachment (sFilePath)

    End If
    
    lobj_cdomsg.send
    Set lobj_cdomsg = Nothing
    SendMail = "ok"
    Exit Function
          
SendMail_Error:
    SendMail = Err.Description

End Function

Public Function SendMailAdjunto(sTo As String, _
                                sSubject As String, _
                                sFrom As String, _
                                sBody As String, _
                                sSmtpServer As String, _
                                iSmtpPort As Integer, _
                                sSmtpUser As String, _
                                sSmtpPword As String, _
                                sFilePath As String, _
                                bSmtpSSL As Boolean, _
                                txtselecciona As String, _
                                txthtml1 As String, _
                                sFilePath2 As String) As String

    Dim txthtml As String

    Dim fso, ts

    On Error GoTo SendMailAdjunto_Error:

    Dim lobj_cdomsg As CDO.Message

    Set lobj_cdomsg = New CDO.Message

    'MsgBox txtselecciona
    If txtselecciona = "H" Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set ts = fso.OpenTextFile(txthtml1)
        txthtml = ts.ReadAll

    End If
    
    lobj_cdomsg.Configuration.Fields(cdoSMTPServer) = sSmtpServer
    lobj_cdomsg.Configuration.Fields(cdoSMTPServerPort) = iSmtpPort
    lobj_cdomsg.Configuration.Fields(cdoSMTPUseSSL) = bSmtpSSL
    lobj_cdomsg.Configuration.Fields(cdoSMTPAuthenticate) = cdoBasic
    lobj_cdomsg.Configuration.Fields(cdoSendUserName) = sSmtpUser
    lobj_cdomsg.Configuration.Fields(cdoSendPassword) = sSmtpPword
    lobj_cdomsg.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
    lobj_cdomsg.Configuration.Fields(cdoSendUsingMethod) = cdoSendUsingPort
    lobj_cdomsg.Configuration.Fields.Update
    lobj_cdomsg.To = sTo
    lobj_cdomsg.From = sFrom
    lobj_cdomsg.Subject = sSubject

    If txtselecciona = "T" Then
        lobj_cdomsg.TextBody = sBody

    End If

    If txtselecciona = "H" Then
        lobj_cdomsg.HTMLBody = txthtml

    End If

    If Trim$(sFilePath) <> vbNullString Then
        lobj_cdomsg.AddAttachment (sFilePath)
        lobj_cdomsg.AddAttachment (sFilePath2)

    End If
    
    lobj_cdomsg.send
    Set lobj_cdomsg = Nothing
    SendMailAdjunto = "ok"
    Exit Function
          
SendMailAdjunto_Error:
    SendMailAdjunto = Err.Description

End Function

Function Imprime_archivojj(path As String, _
                           xsw As Integer, _
                           xtipoletra As String, _
                           nombrefont As String, _
                           xx As String, _
                           letrainterna As String)

    Dim free_file As Integer

    Dim datos     As String

    Dim pos       As Integer

    Dim L         As String

    Dim I         As Integer

    Dim antfont

    Dim Palabra As String

    Dim vbcrlf  As String

    Dim buf     As String

    Dim found   As Integer

    On Error GoTo cmd89000_err

    'Dim p As New RAWPrinter
    '   If Len(Trim(nombrefont)) > 0 Then
    '      ponerfont nombrefont
    '     Else
    '    ponerfont "Courier New"
    'Printer.FontName = "Courier New"
    ' End If
    'If Val(xtipoletra) < 7 Then
    'xtipoletra = "9"
    'End If
    'Printer.FontSize = CInt(xtipoletra)
    'Printer.FontSize = 13
    'p.NewDoc ("My Document")
    'p.PrintText "This is a test."
    'p.PrintFile (path) 'Send file directly to the printer.
    'p.EndDoc
    'Exit Function
    'prueba_imprime path
    'Exit Function
    'found = imprime_archivotexto(path)
    'found = imprime_archivoj(path, 0, xtipoletra)
    'Exit Function
    '---------------------------------------------------------
    'Iniciar_Impresion (path)
    'Finalizar_Impresion
    'Exit Function
    '---------------------------------------------------------
    'Printer.Print
    'Printer.FontBold = False
    'MsgBox Trim(nombrefont)
    If Len(Trim(nombrefont)) > 0 Then
        ponerfont nombrefont
    Else
        ponerfont "Courier New"

        'Printer.FontName = "Courier New"
    End If

    If xx = "N" Then
        Printer.FontBold = False
    Else
        Printer.FontBold = True

    End If

    If Val(xtipoletra) < 7 Then
        xtipoletra = "9"

    End If

    Printer.FontSize = CInt(xtipoletra)
         
    found = imprime_archivojp(path, 0, letrainterna)
        
    Exit Function
    free_file = FreeFile
    vbcrlf = Chr$(10) + Chr$(13)
    Open path For Input As free_file
    datos = input(LOF(free_file), free_file)
    Close free_file

    '----------------------------------
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
            'MsgBox L
            If L = "@" Then
                'MsgBox "abc"
                antfont = Printer.FontSize
                Printer.FontSize = 13

            End If

            If L = "+" Then
                Printer.FontSize = antfont

            End If

            pos = InStr(L, " ")

            If pos = 0 Then
                Palabra = L
                L = ""
            Else
                Palabra = Left$(L, pos)
                L = Mid$(L, pos + 1)

            End If
      
            ' verifica que no se pase del ancho de la hoja
            'If (Printer.CurrentX + Printer.TextWidth(Palabra)) <= Printer.ScaleWidth Then
            ' imprime la palabra
            'MsgBox Palabra
            'If L = "@" Then
            '   antfont = Printer.FontSize
            '   Printer.FontSize = 13
            'End If
            Printer.Print Palabra;
            'If L = "+" Then
            '   Printer.FontSize = antfont
            'End If
        
            ' si no imprime en la siguiente linea
            'Else
            '    Printer.Print
            ' verifica que no se pase del alto de la hoja
            '    If (Printer.CurrentY + Printer.Font.Size) > Printer.ScaleHeight Then
            ' nueva hoja
            '        Printer.NewPage
            '    End If
            ' imprime la palabra
            '    Printer.Print Palabra;
            'End If
        Loop
        Printer.Print
    Loop
          
    ' Fin. Manda a imprimir
    Printer.Print
    Printer.EndDoc
    '----------------------------------
    Exit Function
cmd89000_err:
    MsgBox "Aviso en imprimir_archivo JJ ,no es el Driver correcto Impresora " + error$, 48, "Aviso"
    Exit Function
  
End Function

Function imprime_archivojp(xbuf As String, xsw As Integer, letrainterna As String)

    Dim I         As Long

    Dim max       As Long

    Dim buf       As String

    Dim free_file As Integer

    Dim vr

    Dim antfont

    '' 10/07/2018 Edicion Comanda
    Dim tamcomanda As String

    tamcomanda = busca_TamañoComanda
    '' 10/07/2018 Edicion Comanda

    On Error GoTo cmd9876_err

    antfont = Printer.FontSize
    free_file = FreeFile
    Open xbuf For Input As free_file
    max = LOF(free_file)   'numero de letras

    If max < 1 Then
        Close free_file
        Exit Function

    End If
   
    For I = 1 To max
        vr = DoEvents()

        If opcion3 = "1" Then
            If MsgBox("DESEA CANCELAR LA IMPRESION", 1, "AVISO") = 1 Then
                Close free_file
                Printer.EndDoc
                Exit Function

            End If

            opcion3 = "0"

        End If

        Seek free_file, I
        buf = input$(1, free_file)

        'MsgBox buf
        If buf = "£" Then

            'MsgBox "HOLA"
            If Val(letrainterna) > 6 Then
                antfont = Printer.FontSize
                Printer.FontSize = Val(letrainterna)

            End If

        End If

        'Printer.Print Palabra;
        If buf = "Ø" Then
            Printer.FontSize = antfont

        End If
        
        If buf = Chr(12) Then
            If xsw = 0 Then
                Printer.NewPage

            End If

            GoTo a2

        End If

        If buf <> Chr(12) Then
 
            If buf = "Ñ" Then
                buf = "N"

            End If
        
            '' 10/07/2018 Tamaño grande cabecera comanda
            '''11/08/2017 kenyo Tamaño grande encabezado comanda
                                
            If I < 30 Then
                '' 10/07/2018 Tamaño grande cabecera comanda
                Printer.FontSize = CInt(20)
            Else
                Printer.FontSize = CInt(antfont)

            End If
            
            If buf = "-" Then
                Printer.FontSize = CInt(14)

            End If

            Dim j As Integer

            Printer.FontSize = CInt(antfont)

            ' 10/07/2018 Edicion Comanda
            ' Tamaño 12/14/20/ Descripcion de Producto
            'If buf = "+" Then
            '  j = I
            '  Printer.FontSize = CInt(tamcomanda)
            'End If

            If j <> 0 Then
                If I >= j Then
                    Printer.FontSize = CInt(tamcomanda)

                End If

            End If

            ' 10/07/2018 Edicion Comanda
            
            '''11/08/2017 kenyo Tamaño grande encabezado comanda
            If VBA.UCase(buf) <> buf Then
                Printer.FontSize = CInt(antfont)

            End If

            '''11/08/2017 kenyo Tamaño grande encabezado comanda
            
            If I <= 21 Then
                Printer.FontSize = CInt(14)

            End If

            '             If I = "2" And buf = "a" Then
            '               Printer.FontSize = CInt(18)
            '            End If
            '             If I = "3" And buf = "l" Then
            '               Printer.FontSize = CInt(18)
            '            End If
            '             If I = "4" And buf = "o" Then
            '               Printer.FontSize = CInt(18)
            '            End If
            '             If I = "10" And buf = "n" Then
            '               Printer.FontSize = CInt(18)
            '            End If
            '            If I = "11" And buf = ":" Then
            '               Printer.FontSize = CInt(18)
            '            End If
            '
            '
            '            If I = "14" Then
            '                 If buf <> "L" Then Printer.FontSize = CInt(18)
            '                 If buf = "A" Then
            '                  Printer.FontSize = CInt(antfont)
            '                 End If
            '            End If
            '
            '
            '             If I = "15" Then
            '                If buf <> "I" Then Printer.FontSize = CInt(18)
            '                 If buf = "A" Then
            '                  Printer.FontSize = CInt(antfont)
            '                 End If
            '            End If
            '
            '                ' MESA
            '            If I = "21" And buf = "M" Then
            '               Printer.FontSize = CInt(18)
            '            End If
            '             If I = "22" And buf = "e" Then
            '               Printer.FontSize = CInt(18)
            '            End If
            '             If I = "23" And buf = "s" Then
            '               Printer.FontSize = CInt(18)
            '            End If
            '             If I = "24" And buf = "a" Then
            '               Printer.FontSize = CInt(18)
            '            End If
            '             If I = "25" And buf = ":" Then
            '               Printer.FontSize = CInt(18)
            '            End If
            '
            '            If I = "26" Then
            '                If buf <> "*" Then Printer.FontSize = CInt(18)
            '                If buf = "l" Then
            '                  Printer.FontSize = CInt(antfont)
            '                 End If
            '            End If
            '
            '            If I = "27" Then
            '               If buf <> "*" Then Printer.FontSize = CInt(18)
            '                  If buf = "o" Then
            '                  Printer.FontSize = CInt(antfont)
            '                 End If
            '            End If
            '''11/08/2017 kenyo Tamaño grande encabezado comanda
           
            ''' 14/09/2017  Testing
            ' MESA
            '            If I = "24" And buf = "S" Then
            '               Printer.FontSize = CInt(antfont)
            '            End If
            '             If I = "25" And buf = "a" Then
            '               Printer.FontSize = CInt(antfont)
            '            End If
            '             If I = "26" And buf = "l" Then
            '               Printer.FontSize = CInt(antfont)
            '            End If
            '             If I = "27" And buf = "o" Then
            '               Printer.FontSize = CInt(antfont)
            '            End If
            '             If I = "28" And buf = "n" Then
            '               Printer.FontSize = CInt(antfont)
            '            End If
            '            If I = "29" And buf = ":" Then
            '               Printer.FontSize = CInt(antfont)
            '            End If
            '           '''14/09/2017  Testing
            '
            '        '''11/08/2017 kenyo Tamaño grande encabezado comanda
           
            If Chr(10) <> buf Then
                If xsw = 0 Then
                    If buf <> "£" And buf <> "Ø" Then
                        Printer.Print buf;

                    End If

                End If

            End If

        End If

a2:
    Next I

    Close free_file
    Printer.Print
    Printer.EndDoc
    Exit Function
cmd9876_err:

    'inicio 10/02/2018 pll
    Select Case Err.Number

        Case 482
            Printer.EndDoc

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

    'MsgBox "Error en imprime archivo j " & error$, 48, "Aviso"
    'fin 10/02/2018 pll
    Exit Function

End Function

Function proceso_formatoso(mytablex As ADODB.Recordset, _
                           archivo_formato As String, _
                           archivo_orden As String, _
                           ubicacioni As String, _
                           ubicacionf As String)

    Dim I           As Integer

    Dim j           As Integer

    Dim jj          As Integer

    Dim alibaba     As Integer

    Dim sw1         As Integer

    Dim sw          As Integer

    Dim xxsw        As Integer

    Dim buff        As String

    Dim linea       As String

    Dim campo       As String

    Dim nombrearch  As String   'destino

    Dim nombrearch1 As String  'fuente formato

    Dim variable    As String

    Dim posicioni   As Integer

    Dim posicionf   As Integer

    Dim valor       As String

    Dim found       As Integer

    Dim CAMPO1      As String

    Dim CAMPO2      As String

    Dim campo3      As String

    Dim campo4      As Integer

    Dim campoz      As String

    Dim campoy      As String

    Dim buf         As String

    Dim xcampo      As String

    'MsgBox "abcd"
    On Error GoTo cmd778344_err

    nombrearch1 = archivo_formato
    nombrearch = archivo_orden
    'MsgBox nombrearch1
    'Exit Function
    cerrar_archivo
    Open nombrearch For Append As #1
    Open nombrearch1 For Input As #2
iiniciado:
    xxsw = 0
    sw = 0
    sw1 = 0
    Do
        alibaba = 0

        If EOF(2) Then Exit Do
        Line Input #2, buff

        On Error GoTo 0

        linea = Mid$(buff, 1, Len(buff))

        If Mid$(linea, 1, 1) = ubicacioni Then
            sw1 = 1

        End If

        If Mid$(linea, 1, 1) = ubicacionf Then
            sw1 = 0
            GoTo iiniciado

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
                    found = formateaa("", 1, 2, 0)
                    GoTo iiniciado

                End If

                If sw = 0 And Mid$(linea, j, 1) <> "[" And Mid$(linea, j, 1) <> "]" And Mid$(linea, j, 1) <> "{" And Mid$(linea, j, 1) <> "}" And Mid$(linea, j, 1) <> "/" And Mid$(linea, j, 1) <> "\" And Mid$(linea, j, 1) <> "<" And Mid$(linea, j, 1) <> ">" And Mid$(linea, j, 1) <> "^" And Mid$(linea, j, 1) <> "&" And Mid$(linea, j, 1) <> "$" And Mid$(linea, j, 1) <> "?" Then
                    variable = Mid$(linea, j, 1)
                    found = formateaa(variable, 1, 0, 0)

                End If

                If Mid$(linea, j, 1) = "[" Then
                    sw = 1
                    posicioni = j + 1

                End If

                If sw = 1 And Mid$(linea, j, 1) = "]" Then
                    posicionf = j - 1
                    campo = Mid$(linea, posicioni, posicionf - posicioni + 1)
                    xcampo = campo

                    If InStr(xcampo, ">") > 0 Then  'para tomar de otra base de datos
                        xcampo = campo
                        jj = InStr(xcampo, ">")
                        campoz = Mid$(xcampo, 1, jj - 1)
                        campoy = Mid$(xcampo, jj + 1, Len(xcampo) - (jj))

                        'MsgBox campoz & " " & campoy
                        If campoz = "DELIVERI" Then
                            imprime_susdelivery mytablex, campoy

                        End If

                        If campoz = "FACTURA" Then
                            imprime_susfactura mytablex, campoy

                        End If

                        GoTo ami

                    End If
                    
                    If UCase$(campo) = "COMBO" Then
                        imprime_combinacion Trim("" & mytablex.Fields("producto"))
                        GoTo ami

                    End If

                    If UCase$(campo) = "COMENTARIO" Then
                        imprime_comentariobd mytablex
                        GoTo ami

                    End If

                    'MsgBox campo
                    found = extraer_campos(campo, CAMPO1, CAMPO2, campo3, campo4, ",")
                    buf = Mid$(Trim("" & mytablex.Fields(CAMPO1)), Val(CAMPO2), Val(campo3))
                    found = formateaa(buf, Val(campo3), 0, 0)
ami:

                    alibaba = 0
                    sw = 0
                    posicioni = 0
                    posicionf = 0

                End If

            Next j

            found = formateaa("", 1, 2, 0)

        End If

        '-------------------------
    Loop
    Close #2
    Close #1
    cerrar_archivo
    Exit Function
cmd778344_err:
    MsgBox "Aviso en proceso formatoso " + error$, 48, "Aviso"
    cerrar_archivo
    Exit Function

End Function

Sub imprime_susclientes(mytabley As ADODB.Recordset, campoyy As String)

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    Dim found    As Integer

    Dim CAMPO1   As String

    Dim CAMPO2   As String

    Dim campo3   As String

    Dim campo4   As Integer

    Dim campoz   As String

    Dim campoy   As String

    campoy = campoyy
    mytablex.Open "select * from clientes where codigo='" & Trim("" & mytabley.Fields("codigo")) & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        found = extraer_campos(campoy, CAMPO1, CAMPO2, campo3, campo4, ",")
        buf = "" & mytablex.Fields(CAMPO1)
        found = formateaa(buf, Val(campo3), 0, 0)

    End If

    mytablex.Close

End Sub

Sub imprime_susfactura(mytabley As ADODB.Recordset, campoyy As String)

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    Dim found    As Integer

    Dim CAMPO1   As String

    Dim CAMPO2   As String

    Dim campo3   As String

    Dim campo4   As Integer

    Dim campoz   As String

    Dim campoy   As String

    campoy = campoyy
    'MsgBox campoy
    mytablex.Open "select * from factura where local='" & Trim("" & mytabley.Fields("local")) & "' and tipo='" & Trim("" & mytabley.Fields("tipo")) & "' and serie='" & Trim("" & mytabley.Fields("serie")) & "' and numero='" & Trim("" & mytabley.Fields("numero")) & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        found = extraer_campos(campoy, CAMPO1, CAMPO2, campo3, campo4, ",")
        'MsgBox CAMPO1
        buf = "" & mytablex.Fields(CAMPO1)
        found = formateaa(buf, Val(campo3), 0, 0)

    End If

    mytablex.Close

End Sub

Sub imprime_susdelivery(mytabley As ADODB.Recordset, campoyy As String)

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    Dim found    As Integer

    Dim CAMPO1   As String

    Dim CAMPO2   As String

    Dim campo3   As String

    Dim campo4   As Integer

    Dim campoz   As String

    Dim campoy   As String

    'Exit Sub
    campoy = campoyy
    'MsgBox campoy
    mytablex.Open "select * from deliveri where codigo='" & Trim("" & mytabley.Fields("codigo")) & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        found = extraer_campos(campoy, CAMPO1, CAMPO2, campo3, campo4, ",")
        'MsgBox CAMPO1
        buf = "" & mytablex.Fields(CAMPO1)
        found = formateaa(buf, Val(campo3), 0, 0)

    End If

    mytablex.Close

End Sub

Sub imprime_comentariobd(mytablex As ADODB.Recordset)

    Dim found As Integer

    found = formateaa("*" & mytablex.Fields("observa1"), 30, 2, 0)

End Sub

Sub imprime_combinacion(buf As String)

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    mytablex.Open "select * from combina where producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    Do

        If mytablex.EOF Then Exit Do
        '----------------------------------------------
        found = formateaa("*" & mytablex.Fields("descripciop"), 30, 2, 0)
        'found = formateaa("" & mytablex.Fields("cantidad"), 3, 2, 0)
        '----------------------------------------------
        mytablex.MoveNext
    Loop
    mytablex.Close

End Sub

'''' 25/07/2018 Delivery y Para Llevar desde mozo
' Mesas Nombre 21/05/2018
Function busca_clienteDelivery_mesa(codigo As String, salon As String)

    Dim mytablex  As New ADODB.Recordset

    Dim tiposalon As String

    Dim tabla     As String

    mytablex.Open "SELECT  tipo FROM salon where salon='" & Trim("" & salon) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        tiposalon = "" & Trim$("" & mytablex.Fields("tipo"))

    End If

    mytablex.Close

    If tiposalon = "C" Then
        Exit Function

    End If

    If tiposalon = "D" Then
        tabla = "deliveri"
    ElseIf tiposalon = "L" Then
        tabla = "clientes"

    End If

    Dim mytabley As New ADODB.Recordset

    mytabley.Open "SELECT  nombre from " & tabla & "  where codigo='" & Trim("" & codigo) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        busca_clienteDelivery_mesa = "" & Trim$("" & mytabley.Fields("nombre"))

    End If

    mytabley.Close

End Function

' Mesas Nombre 21/05/2018
'''' 25/07/2018 Delivery y Para Llevar desde mozo

