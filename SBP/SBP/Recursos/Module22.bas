Attribute VB_Name = "Module22"

Function despacho_orden(bxsalon As String, bxmesa As String, xcaja As String)

    Dim mytablex As New ADODB.Recordset

    'MsgBox bxsalon & " " & bxmesa & " " & xcaja
    'Exit Function
    mytablex.Open "select * from dcomanda where salon='" & "" & bxsalon & "' and mesa='" & bxmesa & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Function

    End If

    ncanal = 1
    Do

        If mytablex.EOF Then Exit Do
        despacho_cabeceras mytablex, xcaja
        mytablex.MoveNext
    Loop
    mytablex.MoveFirst
    'imprime los productos en cada impresora
    Do

        If mytablex.EOF Then Exit Do
        despacho_productos mytablex, xcaja
        mytablex.MoveNext
    Loop
    mytablex.Close

End Function

Sub despacho_cabeceras(mytabley As ADODB.Recordset, xcaja As String)

    Dim found As Integer

    Dim oldprinter

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from impod where producto='" & "" & mytabley.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    Do

        If mytablex.EOF Then Exit Do
        FileName = xcaja
        found = borra_nombre("" & FileName)
        Open FileName For Append As #ncanal
        cabecera_ordenn "" & mytabley.Fields("vendedor"), "", "", ""
        Close #ncanal

        If "" & mytablex.Fields("cola") = "S" Then
            oldprinter = Printer.DeviceName
            selecciona_impresoras ("" & mytablex.Fields("puerto"))
            found = Imprime_archivojj("" & FileName, 0, "" & mytable11.Fields("tamanorden"), "" & mytable11.Fields("nombrefont"), "" & mytable11.Fields("BOLD"), "" & mytable11.Fields("letrainterna"))
            selecciona_impresoras (Trim(oldprinter))
        Else
            '---------lpt com ----------------------------------
            found = star_sp342(Trim("" & mytablex.Fields("puerto")), 0)

            '-------------------------------------------
        End If

        mytablex.MoveNext
    Loop
    mytablex.Close

End Sub

Sub despacho_productos(mytabley As ADODB.Recordset, xcaja As String)

    Dim found As Integer

    Dim oldprinter

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from impod where producto='" & "" & mytabley.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    Do

        If mytablex.EOF Then Exit Do
        FileName = xcaja
        found = borra_nombre("" & FileName)
        Open FileName For Append As #ncanal
        detalle_ordenn mytabley
        Close #ncanal

        If "" & mytablex.Fields("cola") = "S" Then
            oldprinter = Printer.DeviceName
            selecciona_impresoras ("" & mytablex.Fields("puerto"))
            found = Imprime_archivojj("" & FileName, 0, "" & mytable11.Fields("tamanorden"), "" & mytable11.Fields("nombrefont"), "" & mytable11.Fields("BOLD"), "" & mytable11.Fields("letrainterna"))
            selecciona_impresoras (Trim(oldprinter))
        Else
            '---------lpt com ----------------------------------
            found = star_sp342(Trim("" & mytablex.Fields("puerto")), 0)

            '-------------------------------------------
        End If

        mytablex.MoveNext
    Loop
    mytablex.Close

End Sub

Sub detalle_ordenn(mytabley As ADODB.Recordset)

    Dim buf      As String

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd4711_err

    '----- formato nuevo
    buf = "" & mytabley.Fields("cantidad")
    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    'MsgBox ""
    buf = "" & Mid$("" & mytabley.Fields("descripcio"), 1, 31)
    found = formateaa(buf, 31, 0, 0)
    found = formateaa(" ", 1, 2, 0)

    If Len("" & mytabley.Fields("descripcio")) > 21 Then
        buf = "      " & Mid$("" & mytabley.Fields("descripcio"), 32, 31)
        'buf = "" & Mid$("" & mytabley.Fields("descripcio"), 32, 31)
        found = formateaa(buf, 31, 0, 0)
        found = formateaa(" ", 1, 2, 0)

    End If

    'verificar si tiene receta
    If imprimecombo_ve("" & mytabley.Fields("producto")) = 1 Then
        mytablex.Open "SELECT * FROM receta where producto='" & "" & mytabley.Fields("producto") & "'", cn, adOpenDynamic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            Do

                If mytablex.EOF Then Exit Do
                '-------------------------------------------
                found = formateaa("++", 2, 0, 0)
                buf = "" & mytablex.Fields("cantidad")
                found = formateaa(buf, 5, 0, 0)
                found = formateaa("", 1, 0, 0)
                'MsgBox ""
                buf = "" & Mid$("" & mytablex.Fields("descripcio"), 1, 31)
                found = formateaa(buf, 31, 0, 0)
                found = formateaa(" ", 1, 2, 0)

                If Len("" & mytablex.Fields("descripcio")) > 21 Then
                    'buf = "      " & Mid$("" & mytablex.Fields("descripcio"), 32, 31)
                    buf = "" & Mid$("" & mytablex.Fields("descripcio"), 32, 31)
                    found = formateaa(buf, 31, 0, 0)
                    found = formateaa(" ", 1, 2, 0)

                End If

                '-------------------------------------------
                mytablex.MoveNext
            Loop

        End If

        mytablex.Close

    End If

    'found = formateaa("------------------------------------- ", 28, 2, 0)
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

    If Len(Trim("" & mytabley.Fields("observa4"))) > 0 Then
        found = combina_imprime("" & mytabley.Fields("producto"))

    End If

    Exit Sub
cmd4711_err:
    MsgBox "Aviso en imprime detalle orden " + error$, 48, "Aviso"
    Exit Sub
    
End Sub

Sub cabecera_ordenn(buvendedor As String, buf1 As String, buf2 As String, buf3 As String)

    Dim found As Integer

    Dim buf   As String

    Dim btipo As String

    On Error GoTo cmd114111_err

    buf = String(42, "-")
    found = formateaa(buf, 45, 2, 0)

    If Len(buf2) > 0 Then
        found = formateaa("       Numero:" & buf2, 28, 2, 0)

    End If

    buf = "     ORDEN DESPACHO " & comanda
    found = formateaa(buf, 28, 2, 0)
    buf = "     Caja :" & caja & " Turno:" & turno
    found = formateaa(buf, 28, 2, 0)
    buf = "Fecha:" & Format(Now, "dd/mm/yyyy") & " " & "Hora :" & Format(Now, "hh:mm:ss")
    found = formateaa(buf, 28, 2, 0)

    If flag_servicio = "A" Then
        found = formateaa("       *** PARA LLEVAR    ***", 28, 2, 0)
        found = formateaa("Nombre:" + Mid$(buf3, 1, 20), 28, 2, 0)
        buf = "Mozo  :"
        found = formateaa(buf, 8, 0, 0)
        found = busca_mesero_vendedor(buvendedor)

    End If
   
    If flag_servicio = "A" Then
        buf = "VENTA RAPIDA"
        found = formateaa(buf, 28, 2, 0)
      
    End If

    If flag_servicio = "C" Then
        buf = "Salon : " & salon & " Mesa:" & mesa
        found = formateaa(buf, 28, 2, 0)
        buf = "Mozo  :"
        found = formateaa(buf, 8, 0, 0)
        found = busca_mesero_vendedor(buvendedor)

    End If

    If flag_servicio = "D" Then
        found = formateaa("       *** DOMICILIO ***", 28, 2, 0)
        found = formateaa(buf, 28, 2, 0)
        imprime_delivery_cliente "" & codigo

    End If

    If flag_servicio <> "A" And flag_servicio <> "D" And flag_servicio <> "C" Then
        buf = "OTROS SERVICIOS"
        found = formateaa(buf, 28, 2, 0)

    End If

    buf = "///" & xnombre
    found = formateaa(buf, 28, 2, 0)
      
    If buf1 = "***ANULADO***" Then
        found = formateaa("ANULADO", 25, 2, 0)

    End If
   
    buf = String(42, "-")
    found = formateaa(buf, 45, 2, 0)

    found = formateaa("CANT", 6, 0, 0)

    found = formateaa("PRODUCTO ", 21, 0, 0)
    found = formateaa(" ", 1, 2, 0)
 
    buf = String(42, "-")
    found = formateaa(buf, 45, 2, 0)

    Exit Sub
cmd114111_err:
    MsgBox "Mensaje,Error en cabecera Pedido " & error$, 48, "Aviso"
    Exit Sub

End Sub

Function busca_mesero_vendedor(buvendedor As String)

    Dim buf      As String

    Dim found    As Integer

    'MsgBox buvendedor
    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM vendedor where codigo='" & buvendedor & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        found = formateaa("", 1, 2, 0)

    End If

    If mytablex.RecordCount > 0 Then
        buf = "" & mytablex.Fields("nombre")
        found = formateaa(buf, 20, 2, 0)
        busca_mesero_vendedor = 1

    End If

    mytablex.Close
    Exit Function

End Function

Sub imprime_delivery_cliente(buf1 As String)

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    Dim found    As Integer

    mytablex.Open "SELECT * FROM clientes where codigo='" & buf1 & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        'found = formateaa(" *** DOMICILIO ***", 36, 2, 0)
        buf = "Telf:" & "" & mytablex.Fields("codigo")
        found = formateaa(buf, 36, 2, 0)
        buf = "Nomb:" & "" & mytablex.Fields("nombre")
        found = formateaa(buf, 36, 2, 0)
        buf = "Dire:" & "" & mytablex.Fields("direccion")
        found = formateaa(buf, 36, 2, 0)

    End If

    mytablex.Close

End Sub

Function combina_imprime(buf)

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    mytablex.Open "select * from _c" & gusuario & " where producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Function

    End If

    Do

        If mytablex.EOF Then Exit Do
        '----------------------------------------------
        found = formateaa("*" & mytablex.Fields("descripciop"), 10, 0, 0)
        found = formateaa("" & mytablex.Fields("cantidad"), 3, 2, 0)
        '----------------------------------------------
        mytablex.MoveNext
    Loop
    mytablex.Close
 
End Function

Function imprimecombo_ve(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT recetaprn FROM producto where producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        If "" & mytablex.Fields("recetaprn") = "S" Then
            imprimecombo_ve = 1

        End If

    End If

    mytablex.Close

End Function

