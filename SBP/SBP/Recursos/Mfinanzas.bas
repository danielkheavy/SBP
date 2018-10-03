Attribute VB_Name = "Mfinanzas"

Type struc_cuentac

    codigo                        As String
    codigo1                        As String
    nombre                        As String
    zona                          As String
    Grupo                         As String
    tipo                          As String
    serie                         As String
    Numero                        As String
    fecha                         As String
    cuota                         As String
    vendedor                      As String
    total                         As String
    abono                         As Double
    interes                       As Double
    saldo                         As Double
    dias                          As String
    descripcio                    As String
    unidad                        As String
    factor                        As String
    cantidad                      As Double
    precio                        As Double
    totalP                        As Double
    moneda                        As String

End Type

Global my_struc_cuentac() As struc_cuentac

Public Sub CuentaXCobrar(my_struc_cuentac() As struc_cuentac, _
                         xcuentaco As String, _
                         fechai As String, _
                         fechaf As String, _
                         tipofecha As String, _
                         local1 As String, _
                         tipo As String, _
                         serie As String, _
                         Numero As String, _
                         codigo As String, _
                         nombre As String, _
                         moneda As String, _
                         vendedor As String, _
                         xtipo As String, _
                         tiposaldo As String, _
                         Combo1, _
                         salida As Boolean, _
                         k As Integer, _
                         valor As String, _
                         my_codcliente As String)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    Dim mytable  As New ADODB.Recordset

    If Len(fechai) <> 10 Then Exit Sub
    If Len(fechaf) <> 10 Then Exit Sub
    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub

    'On Error GoTo CuentaXCobrar

    ReDim my_struc_cuentac(0)

    mysql = "select *" & Chr$(10)
    mysql = mysql & "from " & xcuentaco & "  cx" & Chr$(10)
    mysql = mysql & "where " & Chr$(10)

    If tipofecha = "EMISION" Then
        mysql = mysql & "cx.fecha>='" & Format(fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & " and cx.fecha<='" & Format(fechaf, "YYYYMMDD") & "' " & Chr$(10)

    End If

    If tipofecha = "VENCIMIENTO" Then
        mysql = mysql & "  cx.fechav>='" & Format(fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & " and cx.fechav<='" & Format(fechaf, "YYYYMMDD") & "' " & Chr$(10)

    End If

    If local1 <> "%" Then
        mysql = mysql & " and cx.local='" & local1 & "'" & Chr$(10)

    End If

    If tipo <> "%" Then
        mysql = mysql & " and cx.tipo like '" & tipo & "'" & Chr$(10)

    End If

    If serie <> "%" Then
        mysql = mysql & " and cx.serie like '" & serie & "'" & Chr$(10)

    End If

    If Numero <> "%" Then
        mysql = mysql & " and cx.numero like '" & Numero & "'" & Chr$(10)

    End If

    If codigo <> "%" Then
        mysql = mysql & " and cx.codigo like '" & codigo & "'" & Chr$(10)

    End If

    If nombre <> "%" Then
        mysql = mysql & " and cx.nombre like '" & nombre & "'" & Chr$(10)

    End If

    If moneda <> "%" Then
        mysql = mysql & " and cx.moneda like '" & moneda & "'" & Chr$(10)

    End If

    If vendedor <> "%" Then
        mysql = mysql & " and cx.vendedor like '" & vendedor & "'" & Chr$(10)

    End If

    If xtipo = "CREDITO" Then
        mysql = mysql & " and cx.grupo='C'" & Chr$(10)

    End If

    If xtipo = "ANTICIPO DINERO" Then
        mysql = mysql & " and cx.grupo='A'" & Chr$(10)

    End If

    If xtipo = "DEPOSITO BANCO" Then
        mysql = mysql & " and cx.grupo='D'" & Chr$(10)

    End If

    If xtipo = "ORDEN TRABAJO" Then
        mysql = mysql & " and cx.grupo='O'" & Chr$(10)

    End If

    If tiposaldo = "PENDIENTE" Then
        mysql = mysql & " and (cx.saldo>0 or cx.saldo<0)" & Chr$(10)

    End If

    If tiposaldo = "CANCELADO" Then
        mysql = mysql & " and cx.saldo=0" & Chr$(10)

    End If

    If Combo1 = "Codigo" Then
        mysql = mysql & " order by cx.codigo,cx.moneda,cx.grupo,cx.numero,cx.fechav " & Chr$(10)

    End If

    If Combo1 = "Vendedor" Then
        mysql = mysql & " order by cx.Vendedor,cx.moneda,cx.grupo,cx.numero,cx.fechav " & Chr$(10)

    End If

    If Combo1 = "Zona" Then
        mysql = mysql & " order by cx.Zona,cx.cx.moneda,grupo,cx.numero,cx.fechav " & Chr$(10)

    End If

    'End If
    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF Then
        salida = False
        'Exit Sub
    Else
        salida = True
        mytablex.MoveFirst
        k = 0

        Do Until mytablex.EOF

            If k > 0 Then
                ReDim Preserve my_struc_cuentac(UBound(my_struc_cuentac) + 1)

            End If

            If mytablex.Fields("codigo") <> my_codigo Then
                If mytablex.Fields("codigo") <> "" Then
                    my_struc_cuentac(UBound(my_struc_cuentac)).codigo = mytablex.Fields("codigo")
                Else
                    my_struc_cuentac(UBound(my_struc_cuentac)).codigo = ""

                End If

                my_codigo = mytablex.Fields("codigo")

            End If

            'End If
            If mytablex.Fields("codigo1") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).codigo1 = mytablex.Fields("codigo1")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).codigo1 = ""

            End If

            'End If

            If mytablex.Fields("nombre") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).nombre = mytablex.Fields("nombre")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).nombre = ""

            End If

            'End If
 
            If mytablex.Fields("Zona") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).zona = mytablex.Fields("Zona")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).zona = ""

            End If
    
            If mytablex.Fields("grupo") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).Grupo = mytablex.Fields("grupo")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).Grupo = ""

            End If
   
            If mytablex.Fields("tipo") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).tipo = mytablex.Fields("tipo")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).tipo = ""

            End If
   
            If mytablex.Fields("serie") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).serie = mytablex.Fields("serie")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).serie = ""

            End If
   
            If mytablex.Fields("numero") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).Numero = mytablex.Fields("numero")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).Numero = ""

            End If
   
            If mytablex.Fields("fecha") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).fecha = mytablex.Fields("fecha")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).fecha = ""

            End If

            If mytablex.Fields("cuota") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).cuota = mytablex.Fields("cuota")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).cuota = ""

            End If
    
            If mytablex.Fields("vendedor") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).vendedor = mytablex.Fields("vendedor")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).vendedor = ""

            End If
   
            If mytablex.Fields("Total") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).total = mytablex.Fields("Total")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).total = ""

            End If
   
            If mytablex.Fields("abono") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).abono = mytablex.Fields("abono")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).abono = 0

            End If
   
            If mytablex.Fields("interes") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).interes = mytablex.Fields("interes")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).interes = 0

            End If
   
            If mytablex.Fields("saldo") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).saldo = mytablex.Fields("saldo")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).saldo = 0

            End If
   
            If mytablex.Fields("dias") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).dias = mytablex.Fields("dias")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).dias = 0

            End If
   
            If mytablex.Fields("moneda") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).moneda = mytablex.Fields("moneda")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).moneda = 0

            End If
   
            k = k + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

End Sub

Public Sub titulo_cuentaXCobrar(fechai As String, moneda As String, fechaf As String)
           
    On Error GoTo titulo_cuentaXCobrar

    Set objWorksheet = objWorkBook.Worksheets(1)

    objWorksheet.Cells(10, 5).Font.bold = True
    objWorksheet.Cells(10, 5) = "Cuentas Por Cobrar"
 
    'tipo de moneda
    'objWorksheet.Cells(10, 10) = "Moneda:"
    'If moneda = "D" Then
    '  objWorksheet.Cells(10, 11) = "Dolares"
    'Else
    '  objWorksheet.Cells(10, 11) = "Soles"
    'End If
    'fecha a imprimir
 
    objWorksheet.Cells(11, 10).Font.bold = True
    objWorksheet.Cells(11, 10) = "FECHA HOY :"
    objWorksheet.Cells(11, 10) = Format(Now, "dd/mm/yyyy") & "- HORA " + Format(Now, "HH:MM:SS")
 
    'Aqui los margenes cuadrados
    objWorksheet.Range(Cells(12, 1), Cells(12, 14)).Borders.LineStyle = xlContinuous 'Codigo
    objWorksheet.Range(Cells(12, 1), Cells(12, 14)).Interior.color = RGB(215, 215, 215) 'Codigo

    objWorksheet.Cells(12, 1).Font.bold = True
    objWorksheet.Cells(12, 1) = "Cod.Cliente"
    objWorksheet.Range("A12").ColumnWidth = 12
  
    objWorksheet.Cells(12, 2).Font.bold = True
    objWorksheet.Cells(12, 2) = "Razon Social"
    objWorksheet.Range("B12").ColumnWidth = 30

    objWorksheet.Cells(12, 3).Font.bold = True
    objWorksheet.Cells(12, 3) = "Tipo"
    objWorksheet.Range("C12").ColumnWidth = 8
  
    objWorksheet.Cells(12, 4).Font.bold = True
    objWorksheet.Cells(12, 4) = "Serie"
    objWorksheet.Range("D12").ColumnWidth = 8
  
    objWorksheet.Cells(12, 5).Font.bold = True
    objWorksheet.Cells(12, 5) = "Numero"
    objWorksheet.Range("E12").ColumnWidth = 8
  
    objWorksheet.Cells(12, 6).Font.bold = True
    objWorksheet.Cells(12, 6) = "F.Emision"
    objWorksheet.Range("F12").ColumnWidth = 18
  
    objWorksheet.Cells(12, 7).Font.bold = True
    objWorksheet.Cells(12, 7) = "Cuenta"
    objWorksheet.Range("G12 ").ColumnWidth = 12
  
    objWorksheet.Cells(12, 8).Font.bold = True
    objWorksheet.Cells(12, 8) = "Vendedor"
    objWorksheet.Range("H12").ColumnWidth = 12
  
    objWorksheet.Cells(12, 9).Font.bold = True
    objWorksheet.Cells(12, 9) = "Total"
    objWorksheet.Range("I12").ColumnWidth = 12
  
    objWorksheet.Cells(12, 10).Font.bold = True
    objWorksheet.Cells(12, 10) = "Abono"
    objWorksheet.Range("J12").ColumnWidth = 12
 
    objWorksheet.Cells(12, 11).Font.bold = True
    objWorksheet.Cells(12, 11) = "Intereses"
    objWorksheet.Range("K12").ColumnWidth = 12
  
    objWorksheet.Cells(12, 12).Font.bold = True
    objWorksheet.Cells(12, 12) = "Saldo"
    objWorksheet.Range("L12").ColumnWidth = 12
  
    objWorksheet.Cells(12, 13).Font.bold = True
    objWorksheet.Cells(12, 13) = "Dias"
    objWorksheet.Range("M12").ColumnWidth = 12
  
    objWorksheet.Cells(12, 14).Font.bold = True
    objWorksheet.Cells(12, 14) = "Moneda"
    objWorksheet.Range("M12").ColumnWidth = 12
  
    Exit Sub
titulo_cuentaXCobrar:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

End Sub

Public Sub carga_cuentasXCobrar(my_struc_cuentac() As struc_cuentac, k As Integer)

    Dim my_total As Double

    Dim my_saldo As Double

    v = 13
    h = 0

    On Error GoTo carga_cuentasXCobrar

    For I = 0 To k - 1
        objWorksheet.Cells(v, h + 1).Font.Size = 7 'aqui tamaño letra pll
        objWorksheet.Cells(v, h + 1) = "" & my_struc_cuentac(I).codigo
        objWorksheet.Cells(v, h + 2).Font.Size = 7 'aqui tamaño letra pll
        objWorksheet.Cells(v, h + 2) = "" & my_struc_cuentac(I).nombre
        objWorksheet.Cells(v, h + 3).Font.Size = 7 'aqui tamaño letra pll
        objWorksheet.Cells(v, h + 3) = "" & my_struc_cuentac(I).tipo
        objWorksheet.Cells(v, h + 4).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 4) = my_struc_cuentac(I).serie
        objWorksheet.Cells(v, h + 5).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 5) = my_struc_cuentac(I).Numero
        objWorksheet.Cells(v, h + 6).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 6) = my_struc_cuentac(I).fecha
        objWorksheet.Cells(v, h + 7).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 7) = my_struc_cuentac(I).cuota
        objWorksheet.Cells(v, h + 8).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 8) = my_struc_cuentac(I).vendedor
        objWorksheet.Cells(v, h + 9).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 9) = my_struc_cuentac(I).total
        my_total = my_total + Format(Val(my_struc_cuentac(I).total), "0.00")
     
        objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 10) = my_struc_cuentac(I).abono
        objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 11) = my_struc_cuentac(I).interes
        objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 12) = my_struc_cuentac(I).saldo
        my_saldo = my_saldo + Format(Val(my_struc_cuentac(I).saldo), "0.00")
     
        objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 13) = my_struc_cuentac(I).dias
    
        objWorksheet.Cells(v, h + 14).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 14) = my_struc_cuentac(I).moneda
    
        v = v + 1
    Next I
   
    'aqui los totales
    'h = 3
    objWorksheet.Cells(v, h + 8).Font.bold = True
    objWorksheet.Cells(v, h + 8).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 8) = "TOTALES"
   
    objWorksheet.Cells(v, h + 9).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 9) = "" & Format(my_total, "0.00") 'cantidad

    objWorksheet.Cells(v, h + 12).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 12) = "" & Format(my_saldo, "0.00") 'cantidad
   
    Exit Sub
 
carga_cuentasXCobrar:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

    End Select

End Sub

Public Sub titulo_cXCobrar_cliente(fechai As String, moneda As String, fechaf As String)
           
    On Error GoTo titulo_cXCobrar_cliente

    Set objWorksheet = objWorkBook.Worksheets(1)

    objWorksheet.Cells(10, 5).Font.bold = True
    objWorksheet.Cells(10, 5) = "Consolidado Corporativo de Documentos Por Cobrar"
 
    'tipo de moneda
    'objWorksheet.Cells(10, 10) = "Moneda:"
    'If moneda = "D" Then
    '  objWorksheet.Cells(10, 11) = "Dolares"
    'Else
    '  objWorksheet.Cells(10, 11) = "Soles"
    'End If
    'fecha a imprimir
 
    objWorksheet.Cells(11, 10).Font.bold = True
    objWorksheet.Cells(11, 10) = "FECHA HOY :"
    objWorksheet.Cells(11, 10) = Format(Now, "dd/mm/yyyy") & "- HORA " + Format(Now, "HH:MM:SS")
 
    'Aqui los margenes cuadrados
    objWorksheet.Range(Cells(12, 1), Cells(12, 14)).Borders.LineStyle = xlContinuous 'Codigo
    objWorksheet.Range(Cells(12, 1), Cells(12, 14)).Interior.color = RGB(215, 215, 215) 'Codigo

    objWorksheet.Cells(12, 1).Font.bold = True
    objWorksheet.Cells(12, 1) = "Nomb.Corporativo"
    objWorksheet.Range("A12").ColumnWidth = 20
  
    objWorksheet.Cells(12, 2).Font.bold = True
    objWorksheet.Cells(12, 2) = "Razon Social"
    objWorksheet.Range("B12").ColumnWidth = 30

    objWorksheet.Cells(12, 3).Font.bold = True
    objWorksheet.Cells(12, 3) = "Tipo"
    objWorksheet.Range("C12").ColumnWidth = 8
  
    objWorksheet.Cells(12, 4).Font.bold = True
    objWorksheet.Cells(12, 4) = "Serie"
    objWorksheet.Range("D12").ColumnWidth = 8
  
    objWorksheet.Cells(12, 5).Font.bold = True
    objWorksheet.Cells(12, 5) = "Numero"
    objWorksheet.Range("E12").ColumnWidth = 8
  
    objWorksheet.Cells(12, 6).Font.bold = True
    objWorksheet.Cells(12, 6) = "F.Emision"
    objWorksheet.Range("F12").ColumnWidth = 18
  
    objWorksheet.Cells(12, 7).Font.bold = True
    objWorksheet.Cells(12, 7) = "Cuenta"
    objWorksheet.Range("G12 ").ColumnWidth = 12
  
    objWorksheet.Cells(12, 8).Font.bold = True
    objWorksheet.Cells(12, 8) = "Vendedor"
    objWorksheet.Range("H12").ColumnWidth = 12
  
    objWorksheet.Cells(12, 9).Font.bold = True
    objWorksheet.Cells(12, 9) = "Total"
    objWorksheet.Range("I12").ColumnWidth = 12
  
    objWorksheet.Cells(12, 10).Font.bold = True
    objWorksheet.Cells(12, 10) = "Abono"
    objWorksheet.Range("J12").ColumnWidth = 12
 
    objWorksheet.Cells(12, 11).Font.bold = True
    objWorksheet.Cells(12, 11) = "Intereses"
    objWorksheet.Range("K12").ColumnWidth = 12
  
    objWorksheet.Cells(12, 12).Font.bold = True
    objWorksheet.Cells(12, 12) = "Saldo"
    objWorksheet.Range("L12").ColumnWidth = 12
  
    objWorksheet.Cells(12, 13).Font.bold = True
    objWorksheet.Cells(12, 13) = "Dias"
    objWorksheet.Range("M12").ColumnWidth = 12
  
    objWorksheet.Cells(12, 14).Font.bold = True
    objWorksheet.Cells(12, 14) = "Moneda"
    objWorksheet.Range("N12").ColumnWidth = 12
  
    Exit Sub
titulo_cXCobrar_cliente:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

End Sub

'*********
Public Sub carga_cXCobrar_cliente(my_struc_cuentac() As struc_cuentac, k As Integer)

    Dim my_total As Double

    Dim my_saldo As Double

    v = 13
    h = 0

    On Error GoTo carga_cuentasXCobrar

    For I = 0 To k - 1
        objWorksheet.Cells(v, h + 1).Font.Size = 7 'aqui tamaño letra pll
        objWorksheet.Cells(v, h + 1) = "" & my_struc_cuentac(I).codigo
        objWorksheet.Cells(v, h + 2).Font.Size = 7 'aqui tamaño letra pll
        'objWorksheet.Cells(v, h + 2) = "" & my_struc_cuentac(i).codigo1
        objWorksheet.Cells(v, h + 2) = "" & my_struc_cuentac(I).nombre
        objWorksheet.Cells(v, h + 3).Font.Size = 7 'aqui tamaño letra pll
        objWorksheet.Cells(v, h + 3) = "" & my_struc_cuentac(I).tipo
        objWorksheet.Cells(v, h + 4).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 4) = my_struc_cuentac(I).serie
        objWorksheet.Cells(v, h + 5).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 5) = my_struc_cuentac(I).Numero
        objWorksheet.Cells(v, h + 6).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 6) = my_struc_cuentac(I).fecha
        objWorksheet.Cells(v, h + 7).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 7) = my_struc_cuentac(I).cuota
        objWorksheet.Cells(v, h + 8).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 8) = my_struc_cuentac(I).vendedor
        objWorksheet.Cells(v, h + 9).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 9) = my_struc_cuentac(I).total
        my_total = my_total + Format(Val(my_struc_cuentac(I).total), "0.00")
     
        objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 10) = my_struc_cuentac(I).abono
        objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 11) = my_struc_cuentac(I).interes
        objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 12) = my_struc_cuentac(I).saldo
        my_saldo = my_saldo + Format(Val(my_struc_cuentac(I).saldo), "0.00")
     
        objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 13) = my_struc_cuentac(I).dias
    
        objWorksheet.Cells(v, h + 14).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 14) = my_struc_cuentac(I).moneda
    
        v = v + 1
    Next I
   
    'aqui los totales
    'h = 3
    objWorksheet.Cells(v, h + 8).Font.bold = True
    objWorksheet.Cells(v, h + 8).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 8) = "TOTALES"
   
    objWorksheet.Cells(v, h + 9).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 9) = "" & Format(my_total, "0.00") 'cantidad

    objWorksheet.Cells(v, h + 12).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 12) = "" & Format(my_saldo, "0.00") 'cantidad
   
    Exit Sub
 
carga_cuentasXCobrar:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

    End Select

End Sub

Public Sub CXC_producto(my_struc_cuentac() As struc_cuentac, _
                        xcuentaco As String, _
                        fechai As String, _
                        fechaf As String, _
                        tipofecha As String, _
                        local1 As String, _
                        tipo As String, _
                        serie As String, _
                        Numero As String, _
                        codigo As String, _
                        nombre As String, _
                        moneda As String, _
                        vendedor As String, _
                        xtipo As String, _
                        tiposaldo As String, _
                        Combo1, _
                        salida As Boolean, _
                        k As Integer, _
                        my_codcliente As String)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    If Len(fechai) <> 10 Then Exit Sub
    If Len(fechaf) <> 10 Then Exit Sub
    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub

    On Error GoTo CXC_producto

    ReDim my_struc_cuentac(0)

    mysql = "select de.descripcio,de.unidad,de.factor,de.cantidad,de.precio,de.total as totalP,Cx.*" & Chr$(10)
    mysql = mysql & "from " & xcuentaco & "  cx," & Chr$(10)
    mysql = mysql & "detalle de" & Chr$(10)
    mysql = mysql & "where " & Chr$(10)

    If tipofecha = "EMISION" Then
        mysql = mysql & "cx.fecha>='" & Format(fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & " and cx.fecha<='" & Format(fechaf, "YYYYMMDD") & "' " & Chr$(10)

    End If

    If tipofecha = "VENCIMIENTO" Then
        mysql = mysql & "  cx.fechav>='" & Format(fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & " and cx.fechav<='" & Format(fechaf, "YYYYMMDD") & "' " & Chr$(10)
   
    End If

    If local1 <> "%" Then
        mysql = mysql & " and cx.local='" & local1 & "'" & Chr$(10)

    End If

    If tipo <> "%" Then
        mysql = mysql & " and cx.tipo like '" & tipo & "'" & Chr$(10)

    End If

    If serie <> "%" Then
        mysql = mysql & " and cx.serie like '" & serie & "'" & Chr$(10)

    End If

    If Numero <> "%" Then
        mysql = mysql & " and cx.numero like '" & Numero & "'" & Chr$(10)

    End If

    If codigo <> "%" Then
        mysql = mysql & " and cx.codigo like '" & codigo & "'" & Chr$(10)

    End If

    If nombre <> "%" Then
        mysql = mysql & " and cx.nombre like '" & nombre & "'" & Chr$(10)

    End If

    If moneda <> "%" Then
        mysql = mysql & " and cx.moneda like '" & moneda & "'" & Chr$(10)

    End If

    If vendedor <> "%" Then
        mysql = mysql & " and cx.vendedor like '" & vendedor & "'" & Chr$(10)

    End If

    If xtipo = "CREDITO" Then
        mysql = mysql & " and cx.grupo='C'" & Chr$(10)

    End If

    If xtipo = "ANTICIPO DINERO" Then
        mysql = mysql & " and cx.grupo='A'" & Chr$(10)

    End If

    If xtipo = "DEPOSITO BANCO" Then
        mysql = mysql & " and cx.grupo='D'" & Chr$(10)

    End If

    If xtipo = "ORDEN TRABAJO" Then
        mysql = mysql & " and cx.grupo='O'" & Chr$(10)

    End If

    If tiposaldo = "PENDIENTE" Then
        mysql = mysql & " and (cx.saldo>0 or cx.saldo<0)" & Chr$(10)

    End If

    If tiposaldo = "CANCELADO" Then
        mysql = mysql & " and cx.saldo=0" & Chr$(10)

    End If

    mysql = mysql & "and cx.tipo = de.tipo" & Chr$(10)
    mysql = mysql & "and cx.SERIE = de.serie" & Chr$(10)
    mysql = mysql & "and cx.numero = de.numero" & Chr$(10)
    mysql = mysql & "and cx.GRUPO='C'" & Chr$(10)

    If Combo1 = "Codigo" Then
        mysql = mysql & " order by cx.codigo,cx.moneda,cx.grupo,cx.numero,cx.fechav " & Chr$(10)

    End If

    If Combo1 = "Vendedor" Then
        mysql = mysql & " order by cx.Vendedor,cx.moneda,cx.grupo,cx.numero,cx.fechav " & Chr$(10)

    End If

    If Combo1 = "Zona" Then
        mysql = mysql & " order by cx.Zona,cx.moneda,cx.grupo,cx.numero,cx.fechav " & Chr$(10)

    End If

    'MsgBox buf
    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF Then
        salida = False
        Exit Sub
    Else
        salida = True
        mytablex.MoveFirst
        k = 0

        Do Until mytablex.EOF

            If k > 0 Then
                ReDim Preserve my_struc_cuentac(UBound(my_struc_cuentac) + 1)

            End If
     
            If mytablex.Fields("codigo") <> my_codigo Then
                If mytablex.Fields("codigo") <> "" Then
                    my_struc_cuentac(UBound(my_struc_cuentac)).codigo = mytablex.Fields("codigo")
                Else
                    my_struc_cuentac(UBound(my_struc_cuentac)).codigo = ""

                End If

                my_codigo = mytablex.Fields("codigo")

                If valor = "S" Then
                    If mytablex.Fields("codigo1") <> "" Then
                        my_struc_cuentac(UBound(my_struc_cuentac)).codigo1 = mytablex.Fields("codigo1")
                    Else
                        my_struc_cuentac(UBound(my_struc_cuentac)).codigo1 = ""

                    End If

                End If

                If mytablex.Fields("nombre") <> "" Then
                    my_struc_cuentac(UBound(my_struc_cuentac)).nombre = mytablex.Fields("nombre")
                Else
                    my_struc_cuentac(UBound(my_struc_cuentac)).nombre = ""

                End If

            End If
 
            If mytablex.Fields("descripcio") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).descripcio = mytablex.Fields("descripcio")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).descripcio = ""

            End If
   
            If mytablex.Fields("unidad") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).unidad = mytablex.Fields("unidad")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).unidad = ""

            End If
   
            If mytablex.Fields("factor") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).factor = mytablex.Fields("factor")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).factor = ""

            End If
   
            If mytablex.Fields("cantidad") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).cantidad = mytablex.Fields("cantidad")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).cantidad = 0

            End If
   
            If mytablex.Fields("precio") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).precio = mytablex.Fields("precio")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).precio = 0

            End If
   
            If mytablex.Fields("totalP") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).totalP = mytablex.Fields("totalP")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).totalP = 0

            End If
   
            If mytablex.Fields("Zona") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).zona = mytablex.Fields("Zona")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).zona = ""

            End If
    
            If mytablex.Fields("grupo") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).Grupo = mytablex.Fields("grupo")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).Grupo = ""

            End If
   
            If mytablex.Fields("tipo") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).tipo = mytablex.Fields("tipo")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).tipo = ""

            End If
   
            If mytablex.Fields("serie") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).serie = mytablex.Fields("serie")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).serie = ""

            End If
   
            If mytablex.Fields("numero") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).Numero = mytablex.Fields("numero")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).Numero = ""

            End If
   
            If mytablex.Fields("fecha") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).fecha = mytablex.Fields("fecha")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).fecha = ""

            End If

            If mytablex.Fields("cuota") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).cuota = mytablex.Fields("cuota")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).cuota = ""

            End If
    
            If mytablex.Fields("vendedor") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).vendedor = mytablex.Fields("vendedor")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).vendedor = ""

            End If
   
            If mytablex.Fields("Total") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).total = mytablex.Fields("Total")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).total = ""

            End If
   
            If mytablex.Fields("abono") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).abono = mytablex.Fields("abono")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).abono = ""

            End If
   
            If mytablex.Fields("interes") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).interes = mytablex.Fields("interes")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).interes = 0

            End If
   
            If mytablex.Fields("saldo") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).saldo = mytablex.Fields("saldo")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).saldo = 0

            End If
   
            If mytablex.Fields("dias") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).dias = mytablex.Fields("dias")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).dias = 0

            End If
   
            If mytablex.Fields("moneda") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).moneda = mytablex.Fields("moneda")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).moneda = 0

            End If
   
            k = k + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
CXC_producto:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

End Sub

Public Sub titulo_CXC_Producto(fechai As String, moneda As String, fechaf As String)
           
    On Error GoTo titulo_CXC_Producto

    Set objWorksheet = objWorkBook.Worksheets(1)

    objWorksheet.Cells(10, 5).Font.bold = True
    objWorksheet.Cells(10, 5) = "Cuentas Por Cobrar"
 
    'tipo de moneda
    'objWorksheet.Cells(10, 10) = "Moneda:"
    'If moneda = "D" Then
    '  objWorksheet.Cells(10, 11) = "Dolares"
    'Else
    '  objWorksheet.Cells(10, 11) = "Soles"
    'End If
    'fecha a imprimir
 
    objWorksheet.Cells(11, 10).Font.bold = True
    objWorksheet.Cells(11, 10) = "FECHA HOY :"
    objWorksheet.Cells(11, 10) = Format(Now, "dd/mm/yyyy") & "- HORA " + Format(Now, "HH:MM:SS")
 
    'Aqui los margenes cuadrados
    objWorksheet.Range(Cells(12, 1), Cells(12, 20)).Borders.LineStyle = xlContinuous 'Codigo
    objWorksheet.Range(Cells(12, 1), Cells(12, 20)).Interior.color = RGB(215, 215, 215) 'Codigo

    objWorksheet.Cells(12, 1).Font.bold = True
    objWorksheet.Cells(12, 1) = "Cod.Cliente"
    objWorksheet.Range("A12").ColumnWidth = 12
  
    objWorksheet.Cells(12, 2).Font.bold = True
    objWorksheet.Cells(12, 2) = "Razon Social"
    objWorksheet.Range("B12").ColumnWidth = 30

    objWorksheet.Cells(12, 3).Font.bold = True
    objWorksheet.Cells(12, 3) = "Producto"
    objWorksheet.Range("C12").ColumnWidth = 25
  
    objWorksheet.Cells(12, 4).Font.bold = True
    objWorksheet.Cells(12, 4) = "Unidad"
    objWorksheet.Range("D12").ColumnWidth = 7
  
    objWorksheet.Cells(12, 5).Font.bold = True
    objWorksheet.Cells(12, 5) = "Factor"
    objWorksheet.Range("E12").ColumnWidth = 7
  
    objWorksheet.Cells(12, 6).Font.bold = True
    objWorksheet.Cells(12, 6) = "Cantidad"
    objWorksheet.Range("F12").ColumnWidth = 8
  
    objWorksheet.Cells(12, 7).Font.bold = True
    objWorksheet.Cells(12, 7) = "Precio"
    objWorksheet.Range("G12").ColumnWidth = 8
  
    objWorksheet.Cells(12, 8).Font.bold = True
    objWorksheet.Cells(12, 8) = "Total"
    objWorksheet.Range("H12").ColumnWidth = 8
  
    objWorksheet.Cells(12, 9).Font.bold = True
    objWorksheet.Cells(12, 9) = "Tipo"
    objWorksheet.Range("I12").ColumnWidth = 8
  
    objWorksheet.Cells(12, 10).Font.bold = True
    objWorksheet.Cells(12, 10) = "Serie"
    objWorksheet.Range("J12").ColumnWidth = 8
  
    objWorksheet.Cells(12, 11).Font.bold = True
    objWorksheet.Cells(12, 11) = "Numero"
    objWorksheet.Range("K12").ColumnWidth = 8
  
    objWorksheet.Cells(12, 12).Font.bold = True
    objWorksheet.Cells(12, 12) = "F.Emision"
    objWorksheet.Range("L12").ColumnWidth = 18
  
    objWorksheet.Cells(12, 13).Font.bold = True
    objWorksheet.Cells(12, 13) = "Cuenta"
    objWorksheet.Range("M12 ").ColumnWidth = 12
  
    objWorksheet.Cells(12, 14).Font.bold = True
    objWorksheet.Cells(12, 14) = "Vendedor"
    objWorksheet.Range("N12").ColumnWidth = 12
  
    objWorksheet.Cells(12, 15).Font.bold = True
    objWorksheet.Cells(12, 15) = "Total"
    objWorksheet.Range("L12").ColumnWidth = 12
  
    objWorksheet.Cells(12, 16).Font.bold = True
    objWorksheet.Cells(12, 16) = "Abono"
    objWorksheet.Range("O12").ColumnWidth = 12
 
    objWorksheet.Cells(12, 17).Font.bold = True
    objWorksheet.Cells(12, 17) = "Intereses"
    objWorksheet.Range("P12").ColumnWidth = 12
  
    objWorksheet.Cells(12, 18).Font.bold = True
    objWorksheet.Cells(12, 18) = "Saldo"
    objWorksheet.Range("Q2").ColumnWidth = 12
  
    objWorksheet.Cells(12, 19).Font.bold = True
    objWorksheet.Cells(12, 19) = "Dias"
    objWorksheet.Range("R12").ColumnWidth = 12
  
    objWorksheet.Cells(12, 20).Font.bold = True
    objWorksheet.Cells(12, 20) = "Moneda"
    objWorksheet.Range("S12").ColumnWidth = 12
  
    Exit Sub
titulo_CXC_Producto:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

End Sub

Public Sub carga_CXC_producto(my_struc_cuentac() As struc_cuentac, k As Integer)

    Dim my_total  As Double

    Dim my_saldo  As Double

    Dim my_totalP As Double

    v = 13
    h = 0

    On Error GoTo carga_CXC_producto

    For I = 0 To k - 1
        objWorksheet.Cells(v, h + 1).Font.Size = 7 'aqui tamaño letra pll
        objWorksheet.Cells(v, h + 1) = "" & my_struc_cuentac(I).codigo
        objWorksheet.Cells(v, h + 2).Font.Size = 7 'aqui tamaño letra pll
        objWorksheet.Cells(v, h + 2) = "" & my_struc_cuentac(I).nombre
        objWorksheet.Cells(v, h + 3).Font.Size = 7 'aqui tamaño letra pll
        objWorksheet.Cells(v, h + 3) = "" & my_struc_cuentac(I).descripcio
        objWorksheet.Cells(v, h + 4).Font.Size = 7 'aqui tamaño letra pll
        objWorksheet.Cells(v, h + 4) = "" & my_struc_cuentac(I).unidad
        objWorksheet.Cells(v, h + 5).Font.Size = 7 'aqui tamaño letra pll
        objWorksheet.Cells(v, h + 5) = "" & my_struc_cuentac(I).factor
        objWorksheet.Cells(v, h + 6).Font.Size = 7 'aqui tamaño letra pll
        objWorksheet.Cells(v, h + 6) = "" & my_struc_cuentac(I).cantidad
        objWorksheet.Cells(v, h + 7).Font.Size = 7 'aqui tamaño letra pll
        objWorksheet.Cells(v, h + 7) = "" & my_struc_cuentac(I).precio
        objWorksheet.Cells(v, h + 8).Font.Size = 7 'aqui tamaño letra pll
        objWorksheet.Cells(v, h + 8) = "" & my_struc_cuentac(I).totalP
        my_totalP = my_totalP + Format(Val(my_struc_cuentac(I).totalP), "0.00")

        objWorksheet.Cells(v, h + 9).Font.Size = 7 'aqui tamaño letra pll
        objWorksheet.Cells(v, h + 9) = "" & my_struc_cuentac(I).tipo
        objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 10) = my_struc_cuentac(I).serie
        objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 11) = my_struc_cuentac(I).Numero
        objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 12) = my_struc_cuentac(I).fecha
        objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 13) = my_struc_cuentac(I).cuota
        objWorksheet.Cells(v, h + 14).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 14) = my_struc_cuentac(I).vendedor
        objWorksheet.Cells(v, h + 15).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 15) = my_struc_cuentac(I).total
        my_total = my_total + Format(Val(my_struc_cuentac(I).total), "0.00")
     
        objWorksheet.Cells(v, h + 16).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 16) = my_struc_cuentac(I).abono
        objWorksheet.Cells(v, h + 17).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 17) = my_struc_cuentac(I).interes
        objWorksheet.Cells(v, h + 18).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 18) = my_struc_cuentac(I).saldo
        my_saldo = my_saldo + Format(Val(my_struc_cuentac(I).saldo), "0.00")
     
        objWorksheet.Cells(v, h + 19).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 19) = my_struc_cuentac(I).dias
    
        objWorksheet.Cells(v, h + 20).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 20) = my_struc_cuentac(I).moneda
    
        v = v + 1
    Next I
   
    'aqui los totales
    'h = 3
    objWorksheet.Cells(v, h + 7).Font.bold = True
    objWorksheet.Cells(v, h + 7).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 7) = "TOTALES"
   
    objWorksheet.Cells(v, h + 8).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 8) = "" & Format(my_totalP, "0.00") 'cantidad
   
    objWorksheet.Cells(v, h + 15).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 15) = "" & Format(my_total, "0.00") 'cantidad

    objWorksheet.Cells(v, h + 14).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 166) = "" & Format(my_saldo, "0.00") 'cantidad
   
    Exit Sub
 
carga_CXC_producto:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

    End Select

End Sub

Public Sub DocuXCobrar(my_struc_cuentac() As struc_cuentac, _
                       xcuentaco As String, _
                       fechai As String, _
                       fechaf As String, _
                       tipofecha As String, _
                       local1 As String, _
                       tipo As String, _
                       serie As String, _
                       Numero As String, _
                       codigo As String, _
                       nombre As String, _
                       moneda As String, _
                       vendedor As String, _
                       xtipo As String, _
                       tiposaldo As String, _
                       Combo1, _
                       salida As Boolean, _
                       k As Integer, _
                       my_codcliente As String)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    Dim mytable  As New ADODB.Recordset

    If Len(fechai) <> 10 Then Exit Sub
    If Len(fechaf) <> 10 Then Exit Sub
    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub

    ReDim my_struc_cuentac(0)

    mysql = "select distinct codigo1" & Chr$(10)
    mysql = mysql & "From clientes " & Chr$(10)
    mysql = mysql & "where CODIGO ='" & my_codcliente & "'" & Chr$(10)
   
    mytable.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytable.RecordCount > 0 Then
        my_codigo1 = mytable.Fields("codigo1")

    End If

    mytable.Close
    '***
    mysql = ""
    mysql = "SELECT t.nombre as corporativo,a.codigo,a.codigo1,b.tipo,b.SERIE,b.NUMERO,a.nombre," & Chr$(10)
    mysql = mysql & "b.fecha,b.zona,b.grupo,b.cuota,b.vendedor,b.total," & Chr$(10)
    mysql = mysql & "b.abono , b.interes, b.saldo, b.dias,b.moneda" & Chr$(10)
    mysql = mysql & "From" & Chr$(10)
    mysql = mysql & "(select nombre from clientes" & Chr$(10)
    mysql = mysql & "where codigo ='" & my_codcliente & "')t," & Chr$(10)
    mysql = mysql & "(select codigo,codigo1,nombre" & Chr$(10)
    mysql = mysql & "From clientes" & Chr$(10)
    mysql = mysql & "where codigo1='" & my_codigo1 & "')a," & Chr$(10)
    mysql = mysql & "(select c.TIPO,c.serie,c.numero," & Chr$(10)
    mysql = mysql & "c.fecha,c.fechav,c.zona,c.grupo,c.cuota,c.vendedor,c.total," & Chr$(10)
    mysql = mysql & "c.abono , c.interes, c.saldo, c.dias,c.moneda" & Chr$(10)
    mysql = mysql & "from factura f," & Chr$(10)
    mysql = mysql & "cuentac c" & Chr$(10)
    mysql = mysql & "where f.codigo='" & my_codcliente & "'" & Chr$(10)
    mysql = mysql & "and f.fecha>='" & Format(fechai, "YYYYMMDD") & "'" & Chr$(10)
    mysql = mysql & "and f.fecha<='" & Format(fechaf, "YYYYMMDD") & "' " & Chr$(10)

    '''
    If local1 <> "%" Then
        mysql = mysql & " and c.local='" & local1 & "'" & Chr$(10)

    End If

    If tipo <> "%" Then
        mysql = mysql & " and c.tipo like '" & tipo & "'" & Chr$(10)

    End If

    If serie <> "%" Then
        mysql = mysql & " and c.serie like '" & serie & "'" & Chr$(10)

    End If

    If Numero <> "%" Then
        mysql = mysql & " and c.numero like '" & Numero & "'" & Chr$(10)

    End If

    If codigo <> "%" Then
        mysql = mysql & " and c.codigo like '" & codigo & "'" & Chr$(10)

    End If

    If nombre <> "%" Then
        mysql = mysql & " and c.nombre like '" & nombre & "'" & Chr$(10)

    End If

    If vendedor <> "%" Then
        mysql = mysql & " and c.vendedor like '" & vendedor & "'" & Chr$(10)

    End If

    If xtipo = "CREDITO" Then
        mysql = mysql & " and c.grupo='C'" & Chr$(10)

    End If

    If xtipo = "ANTICIPO DINERO" Then
        mysql = mysql & " and c.grupo='A'" & Chr$(10)

    End If

    If xtipo = "DEPOSITO BANCO" Then
        mysql = mysql & " and c.grupo='D'" & Chr$(10)

    End If

    If xtipo = "ORDEN TRABAJO" Then
        mysql = mysql & " and c.grupo='O'" & Chr$(10)

    End If

    If tiposaldo = "PENDIENTE" Then
        mysql = mysql & " and (c.saldo>0 or c.saldo<0)" & Chr$(10)

    End If

    If tiposaldo = "CANCELADO" Then
        mysql = mysql & " and c.saldo=0" & Chr$(10)

    End If

    mysql = mysql & "and c.codigo = f.codigo" & Chr$(10)
    mysql = mysql & "and c.tipo = f.tipo" & Chr$(10)
    mysql = mysql & "and f.serie= c.serie" & Chr$(10)
    mysql = mysql & "and f.numero = c.numero" & Chr$(10)
    mysql = mysql & "and c.moneda like '" & moneda & "')b" & Chr$(10)
    'mysql = mysql & "order by b.fecha" & Chr$(10)
  
    If Combo1 = "Codigo" Then
        mysql = mysql & " order by a.codigo,b.grupo,b.numero,b.fechav " & Chr$(10)

    End If

    If Combo1 = "Vendedor" Then
        mysql = mysql & " order by b.Vendedor,b.grupo,b.numero,b.fechav " & Chr$(10)

    End If

    If Combo1 = "Zona" Then
        mysql = mysql & " order by b.Zona,b.grupo,b.numero,b.fechav " & Chr$(10)

    End If

    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF Then
        salida = False
        Exit Sub
    Else
        salida = True
        mytablex.MoveFirst
        k = 0

        Do Until mytablex.EOF

            If k > 0 Then
                ReDim Preserve my_struc_cuentac(UBound(my_struc_cuentac) + 1)

            End If

            If mytablex.Fields("corporativo") <> my_codigo Then
                If mytablex.Fields("corporativo") <> "" Then
                    my_struc_cuentac(UBound(my_struc_cuentac)).codigo = mytablex.Fields("corporativo")
                Else
                    my_struc_cuentac(UBound(my_struc_cuentac)).codigo = ""

                End If

                my_codigo = mytablex.Fields("corporativo")
   
                If mytablex.Fields("codigo1") <> "" Then
                    my_struc_cuentac(UBound(my_struc_cuentac)).codigo1 = mytablex.Fields("codigo1")
                Else
                    my_struc_cuentac(UBound(my_struc_cuentac)).codigo1 = ""

                End If

            End If

            If mytablex.Fields("nombre") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).nombre = mytablex.Fields("nombre")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).nombre = ""

            End If
 
            If mytablex.Fields("Zona") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).zona = mytablex.Fields("Zona")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).zona = ""

            End If
    
            If mytablex.Fields("grupo") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).Grupo = mytablex.Fields("grupo")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).Grupo = ""

            End If
   
            If mytablex.Fields("tipo") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).tipo = mytablex.Fields("tipo")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).tipo = ""

            End If
   
            If mytablex.Fields("serie") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).serie = mytablex.Fields("serie")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).serie = ""

            End If
   
            If mytablex.Fields("numero") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).Numero = mytablex.Fields("numero")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).Numero = ""

            End If
   
            If mytablex.Fields("fecha") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).fecha = mytablex.Fields("fecha")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).fecha = ""

            End If

            If mytablex.Fields("cuota") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).cuota = mytablex.Fields("cuota")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).cuota = ""

            End If
    
            If mytablex.Fields("vendedor") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).vendedor = mytablex.Fields("vendedor")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).vendedor = ""

            End If
   
            If mytablex.Fields("Total") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).total = mytablex.Fields("Total")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).total = ""

            End If
   
            If mytablex.Fields("abono") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).abono = mytablex.Fields("abono")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).abono = 0

            End If
   
            If mytablex.Fields("interes") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).interes = mytablex.Fields("interes")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).interes = 0

            End If
   
            If mytablex.Fields("saldo") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).saldo = mytablex.Fields("saldo")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).saldo = 0

            End If
   
            If mytablex.Fields("dias") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).dias = mytablex.Fields("dias")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).dias = 0

            End If
   
            If mytablex.Fields("moneda") <> "" Then
                my_struc_cuentac(UBound(my_struc_cuentac)).moneda = mytablex.Fields("moneda")
            Else
                my_struc_cuentac(UBound(my_struc_cuentac)).moneda = 0

            End If
   
            k = k + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

End Sub

