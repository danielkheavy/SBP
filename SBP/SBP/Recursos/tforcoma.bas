Attribute VB_Name = "Module21"

Public ArchivoAdjunto As String

Sub formato_precuenta(xsalon As String, xmesa As String)

    Dim mytabley As New ADODB.Recordset

    Dim found    As Integer

    Dim I        As Integer

    Dim oldprinter

    Dim archivo_formato As String

    On Error GoTo cmd450009_err

    found = estado_mesas("" & xsalon, "" & xmesa, "2")
    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    found = borra_nombre("" & FileName)
    archivo_formato = "precuenta"

    If Len(archivo_formato) = 0 Then
        MsgBox "No existe archivo formato ", 48, "Aviso"
        Exit Sub

    End If

    mytabley.Open "SELECT * FROM festadocuenta where  salon='" & xsalon & "' and mesa='" & xmesa & "'", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount = 0 Then
        Exit Sub

    End If

    found = proceso_precuenta(archivo_formato, mytabley, "{", "}", xsalon, xmesa)
    mytabley.Close
       
    'detalle de pagos
    mytabley.Open "SELECT * FROM dcomanda where  salon='" & xsalon & "' and mesa='" & xmesa & "'", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount > 0 Then 'si existe
        Do

            If mytabley.EOF Then Exit Do
            If "" & mytabley.Fields("dua") <> "R" Then
                found = proceso_precuenta(archivo_formato, mytabley, "/", "\", xsalon, xmesa)

            End If

            mytabley.MoveNext
        Loop

    End If

    mytabley.Close
       
    mytabley.Open "SELECT * FROM festadocuenta where  salon='" & xsalon & "' and mesa='" & xmesa & "'", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        found = proceso_precuenta(archivo_formato, mytabley, "$", "?", xsalon, xmesa)

    End If

    mytabley.Close
       
    'impresiones
    If "" & mytable11.Fields("eccola") <> "S" Then
        '------------------------------------
        found = star_sp342("" & mytable11.Fields("ecpuerto"), 0)
        found = corte_papel("" & mytable11.Fields("ecpuerto"), 1)

        '------------------------------------
    End If

    If "" & mytable11.Fields("eccola") = "S" Then
        oldprinter = Printer.DeviceName
        selecciona_impresoras ("" & mytable11.Fields("ecpuerto"))
        found = Imprime_archivojj(FileName, 0, "" & mytable11.Fields("tamanoletra"), "" & mytable11.Fields("nombrefont"), "" & mytable11.Fields("BOLD"), "" & mytable11.Fields("letrainterna"))
        selecciona_impresoras (oldprinter)

    End If
       
    Exit Sub
cmd450009_err:
    MsgBox "Aviso Proceso en Compilando " & error$, 24, "Aviso"
    Exit Sub

End Sub

Function proceso_precuenta(archivo_formato As String, _
                           mytablex As ADODB.Recordset, _
                           ubicacioni As String, _
                           ubicacionf As String, _
                           bxsalon As String, _
                           bxmesa As String)

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

                xxsw = 1

                If Mid$(linea, j, 1) = "[" Then
                    sw = 1
                    posicioni = j + 1

                End If

                If sw = 1 And Mid$(linea, j, 1) = "]" Then
                    posicionf = j - 1
                    campo = Mid$(linea, posicioni, posicionf - posicioni + 1)
                    alibaba = 0
                    valor = busca_precuenta("dcomanda", mytablex, campo, bxsalon, bxmesa)
                    sw = 0
                    posicioni = 0
                    posicionf = 0

                    If alibaba = 1 Then
                      
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
    MsgBox "xx.Existe Un error en Proceso Formatos PROCESO PRECUENTA " & error$, 24, "Aviso"
    cerrar_archivo
    Exit Function
error_lectura:
    MsgBox "Error en Proceso_Precuenta " + error$, 24, "Aviso"
    cerrar_archivo
    Exit Function
    
End Function

Function busca_precuenta(tablabasedatos, _
                         mytablex As ADODB.Recordset, _
                         campo As String, _
                         bxsalon As String, _
                         bxmesa As String)

    Dim CAMPO1     As String

    Dim CAMPO2     As String

    Dim campo3     As String

    Dim campo4     As Integer

    Dim ponemoneda As String

    Dim j          As Integer

    Dim campoz     As String

    Dim campoy     As String

    Dim mytabley   As New ADODB.Recordset

    Dim mytablez   As New ADODB.Recordset

    On Error GoTo cmd89busca_err

    Dim buf   As String

    Dim found As Integer

    campo4 = 0
    buf = campo

    campoz = ""
    campoy = ""

    If InStr(buf, ">") > 0 Then  'para tomar de otra base de datos
        'MsgBox buf
        j = InStr(buf, ">")
        campoz = Mid$(buf, 1, j - 1)
        campoy = Mid$(buf, j + 1, Len(buf) - (j))

        'MsgBox campoz
        'MsgBox "" & mytablex.Fields("vendedor")
        If campoz = "VENDEDOR" Then
            'MsgBox "" & mytablex.Fields("MESA")
            mytablez.Open "SELECT vendedor FROM dcomanda where  salon='" & "" & mytablex.Fields("salon") & "' and mesa='" & "" & mytablex.Fields("mesa") & "'", cn, adOpenDynamic, adLockOptimistic

            If mytablez.RecordCount > 0 Then
                mytabley.Open "SELECT * FROM vendedor where  codigo='" & "" & mytablez.Fields("vendedor") & "'", cn, adOpenDynamic, adLockOptimistic

                If mytabley.RecordCount > 0 Then 'si existe
                    'MsgBox "" & mytabley.Fields("CODIGO")
                    '----------------------------
                    found = extraer_campos(campoy, CAMPO1, CAMPO2, campo3, campo4, ",")
                    buf = "" & mytabley.Fields(CAMPO1)
                    found = formateaa(buf, Val(campo3), 0, 0)

                    '----------------------------
                End If

                mytabley.Close

            End If

            mytablez.Close
            Exit Function

        End If

    End If

    If InStr(campo, ",") > 0 Then   'si es comna
        found = extraer_campos(buf, CAMPO1, CAMPO2, campo3, campo4, ",")

        If Mid$(campo, 1, 1) = "@" Then   'Esto numeros a letras
            CAMPO1 = Mid$(CAMPO1, 2, Len(CAMPO1) - 1)
            buf = ""
            buf = pone_letras("" & mytablex.Fields(CAMPO1), "" & mytablex.Fields("moneda"), campo4)
            buf = Mid$(buf, Val(CAMPO2), Val(campo3))
            found = formateaa(buf, Len(buf), 0, 0)
            Exit Function

        End If

    Else    'si es :

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

        If Mid$(campo, 1, 1) = "@" Then
            buf = ""
            CAMPO1 = Mid$(CAMPO1, 2, Len(CAMPO1) - 1)
            buf = pone_letras("" & mytablex.Fields(CAMPO1), "" & mytablex.Fields("moneda"), 0)
            buf = Mid$(buf, Val(CAMPO2), Val(campo3))
            found = formateaa(buf, Len(buf), 0, 0)
            Exit Function

        End If

    End If
   
    If Val(CAMPO2) > 0 And Val(campo3) > 0 Then
        buf = Mid$("" & mytablex.Fields(CAMPO1), Val(CAMPO2), Val(campo3))
    Else
        buf = "" & mytablex.Fields(CAMPO1)

    End If

    'MsgBox CAMPO1 & " " & mytablex.Fields(CAMPO1).Type
    Select Case Val("" & mytablex.Fields(CAMPO1).Type)

        Case 3, 4  'integer

            If campo4 = 1 Then
                buf = Format(Int(Val(buf)), "0")

            End If

            found = formateaa(buf, Val(campo3), 0, 1)

        Case 5  'double

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

        Case 202, 135
            found = formateaa(buf, Val(campo3), 0, 0)

    End Select

    Exit Function
cmd89busca_err:
    MsgBox "Aviso en busca precuenta " + error$, 48, "Aviso"
    Exit Function

End Function

Function sumar_destadocuenta(xsalon As String, xmesa As String)

    Dim buf As String

    On Error GoTo cmd9056_err

    cn.Execute ("delete from festadocuenta where salon='" & xsalon & "' and mesa='" & xmesa & "' ")
    buf = "INSERT INTO festadocuenta "
    buf = buf & " (salon, mesa,neto,subtotal,impuesto,total,descuento)"
    buf = buf & "SELECT     salon,mesa,sum(neto),sum(subtotal),sum(impuesto),sum(total),sum(descuento)"
    buf = buf & " From dcomanda  "
    buf = buf & " WHERE  salon='" & xsalon & "' and mesa='" & xmesa & "'"
    buf = buf & " GROUP BY salon,mesa"
    cn.Execute (buf)
    buf = "update festadocuenta set fecha='" & Format(Now, "YYYYMMDD") & "',"
    buf = buf & "hora='" & Format(Now, "hh:mm:ss") & "',"
    buf = buf & "caja='" & tptovta.caja & "',"
    buf = buf & "turno='" & tptovta.turno & "'"
    cn.Execute (buf)
    sumar_destadocuenta = 1
    Exit Function
cmd9056_err:
    Exit Function

End Function

Function estado_mesas(buf1 As String, buf2 As String, buf3 As String)

    Dim mytablex As New ADODB.Recordset

    'MsgBox buf1 & " " & buf2
    mytablex.Open "SELECT * FROM mesa where salon='" & buf1 & "' and mesa='" & buf2 & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        mytablex.Fields("estado") = buf3
        mytablex.Update

    End If

    mytablex.Close

End Function

Public Sub CreateBackup()

    Dim strSQL    As String

    Dim xcampo1   As String

    Dim xcampo2   As String

    Dim xcampo3   As String

    Dim nombreRar As String

    On Error GoTo errHandler

    If MsgBox("Desea Relizar Copia Seguridad,Demora un Poco debe saber esperar...", 1, "Aviso") <> 1 Then Exit Sub
    'cn.CommandTimeout = 0
    'DoEvents
    found = extraer_campos1(menup.gempresa, xcampo1, xcampo2, xcampo3)
    'MsgBox extra_loquesea1(menup.gempresa)
    strSQL = "BACKUP DATABASE [" & Trim(extra_loquesea1(xcampo2)) & "] TO DISK = '" & globalpath & "\" & Format(Now, "ddmmyy") & ".BAK'"
  
    'MsgBox strSQL
    cn.Execute strSQL
    MsgBox "Proceso Realizado ", 48, "Aviso"
    'cn.CommitTrans
    Exit Sub
errHandler:
    MsgBox "No se pudo crear copia seguridad " + error$, 48, "Aviso"
    Exit Sub

End Sub

Public Sub CreateBackupBd()

    Dim strSQL        As String

    Dim xcampo1       As String

    Dim xcampo2       As String

    Dim xcampo3       As String

    Dim nombreArchivo As String

    On Error GoTo errHandler

    If MsgBox("Desea Relizar Copia Seguridad,Demora un Poco debe saber esperar...", 1, "Aviso") <> 1 Then Exit Sub
 
    found = extraer_campos1(menup.gempresa, xcampo1, xcampo2, xcampo3)
    strSQL = "BACKUP DATABASE [" & Trim(extra_loquesea1(xcampo2)) & "] TO DISK = 'D:\" & basedatos & "" & Format(Now, "ddmmyy") & ".BAK'"
      
    cn.Execute strSQL
    
    MsgBox "Proceso Realizado ", 48, "Aviso"
  
    Exit Sub
errHandler:
    MsgBox "No se pudo crear copia seguridad " + error$, 48, "Aviso"
    Exit Sub

End Sub

Public Sub ComprimeBackupBd()

    'carpetaToExtract = "D:\" & basedatos & Format(Now, "ddmmyy") & ".BAK"
    'ArchivoAdjunto = "D:\" & basedatos & Format(Now, "ddmmyy") & ".rar"

    Dim carpetaToExtract As String

    carpetaToExtract = "D:\" & basedatos & Format(Now, "ddmmyy") & ".BAK"
    ArchivoAdjunto = "D:\" & basedatos & Format(Now, "ddmmyy") & ".rar"
    Shell "C:\Program Files\WinRAR\WinRAR.exe a " & ArchivoAdjunto & " " & carpetaToExtract, vbHide
  
End Sub

Public Sub eliminar_BackupBd()

    Dim fso As New Scripting.FileSystemObject

    On Error GoTo eliminar_BackupBd

    Kill ("D:\" & basedatos & "" & Format(Now, "ddmmyy") & ".bak")

    Dir1.refresh

eliminar_BackupBd:
    Exit Sub

End Sub

Sub envio_correosBackupBd()

    Dim txtserver     As String

    Dim txtusername   As String

    Dim txtpassword   As String

    Dim txtport       As String

    Dim txtto         As String

    Dim chkssl        As String

    Dim txtfromname   As String

    Dim txtfromemail  As String

    Dim txtattach     As String

    Dim txtsubject    As String

    Dim txtmsg        As String

    Dim retval        As String

    Dim txthtml       As String

    Dim txtselecciona As String

    'Dim txtselecciona As String
    Dim mytablex      As New ADODB.Recordset

    Dim buf           As String

    On Error GoTo cmd0905677_err

    mytablex.Open "select * from correos where cosms='11'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        MsgBox "Correo No Configurado", vbCritical, "Message"

    End If

    If mytablex.RecordCount > 0 Then
        txtserver = Trim("" & mytablex.Fields("txtserver"))
        txtusername = Trim("" & mytablex.Fields("txtusername"))
        txtpassword = Trim("" & mytablex.Fields("txtpassword"))
        txtfromname = Trim("" & mytablex.Fields("txtfromname"))
        txtfromemail = Trim("" & mytablex.Fields("txtfromemail"))
        txtport = Trim("" & mytablex.Fields("txtport"))
        txtselecciona = Trim("" & mytablex.Fields("txtselecciona"))
        chkssl = Trim("" & mytablex.Fields("chkssl"))
        txtto = Trim("" & mytablex.Fields("txtfromemail"))

        txtattach = ArchivoAdjunto 'Rar Archivo Adjunto

        txtsubject = "Backup : " + Format(Now, "dd/mm/yyyy")
        txtmsg = Trim("" & mytablex.Fields("txtmsg"))
        txtmsg = txtmsg & Chr$(10) & Chr$(13) & ""
        txtmsg = txtmsg & Format(Now, "dd/mm/yyyy") + " " + Format(Now, "hh:mm:ss")

        If Len(Trim("" & mytablex.Fields("txtfromemail"))) > 0 Then
            txtto = Trim("" & mytablex.Fields("txtfromemail"))
            retval = SendMail(Trim$(txtto), Trim$(txtsubject), Trim$(txtfromname) & "<" & Trim$(txtfromemail) & ">", Trim$(txtmsg), Trim$(txtserver), CInt(Trim$(txtport)), Trim$(txtusername), Trim$(txtpassword), Trim$(txtattach), True, txtselecciona, txthtml)
   
        End If

        MsgBox "Correo Enviado ", 48, "Aviso"

    End If

    mytablex.Close

    Exit Sub
cmd0905677_err:
    MsgBox "No se Pudo enviar Correo... " + error$, 48, "Aviso"
    Exit Sub

End Sub

