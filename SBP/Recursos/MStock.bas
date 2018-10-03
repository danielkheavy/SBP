Attribute VB_Name = "MStock"

Type struc_Stock_SAnterior

    producto                           As String
    cantidad                           As String

End Type

Global my_struc_Stock_SAnterior() As struc_Stock_SAnterior

'inicio 13/07/2017 pll
Public Sub descarga_saldo(xlocal As String, _
                          xtipo As String, _
                          xserie As String, _
                          xnumero As String, _
                          sw As Integer, _
                          tipoarch As String, _
                          xtipo1 As String, _
                          my_nuevaCantidad As Double)

    Dim sdx          As Double

    Dim signo        As Double

    Dim sww          As Integer

    Dim mytablestock As New ADODB.Recordset

    Dim mytablex     As New ADODB.Recordset

    Dim mytabley     As New ADODB.Recordset

    Dim mytable      As New ADODB.Recordset

    Dim buf          As String

    Dim found        As Integer

    Dim mysql        As String

    On Error GoTo cmd19_err

    sww = 0
    'AHORA HAY QUE VALIDAR QUE no existe ya cruzado el documento----
  
    mysql = "SELECT * FROM " & cgusuario & Chr$(10)
    mysql = mysql & "where  local='" & xlocal & "' " & Chr$(10)
    mysql = mysql & "and tipo='" & xtipo & "'" & Chr$(10)
    mysql = mysql & "and serie='" & xserie & "'" & Chr$(10)
    mysql = mysql & "and numero='" & xnumero & "'" & Chr$(10) 'numero ventas
 
    mytablestock.Open mysql, cn, adOpenStatic, adLockOptimistic
 
    If mytablestock.RecordCount > 0 Then  'si existe
        If Len(xtipo1) > 0 Then
            found = ve_descarga(xtipo1)

            If found = 1 Then
                sww = 1

            End If

        End If

    End If

    buf = dgusuariog

    If tipoarch = "1" Then
        buf = "detalle"

    End If

    mysql = ""
    mysql = "SELECT * FROM " & buf & Chr$(10)
    mysql = mysql & "where  local='" & xlocal & "'" & Chr$(10)
    mysql = mysql & "and tipo='" & xtipo & "' " & Chr$(10)
    mysql = mysql & "and serie='" & xserie & "'" & Chr$(10)
    mysql = mysql & "and numero='" & xnumero & "'" & Chr$(10)
 
    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic
  
    If mytablex.RecordCount = 0 Then 'si existe
        'Exit Sub
        'inicio 14/09/2017 pll
        sww = 0
        sw = 2
        mytablex.Close

        If acu = "R" Then
            mysql = ""
            mysql = "SELECT * FROM dcotizav" & Chr$(10)
            mysql = mysql & "where  local='" & xlocal & "'" & Chr$(10)
            mysql = mysql & "and tipo='H' " & Chr$(10)
            mysql = mysql & "and serie='" & my_serie1 & "'" & Chr$(10)
            mysql = mysql & "and numero='" & my_numero1 & "'" & Chr$(10)
 
            mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

            If mytablex.RecordCount = 0 Then 'si existe
                sww = 0
                sw = 2
            Else
                sww = 0
                sw = 1

            End If

        End If

    Else
        sw = 1
        sdx = "" & mytablex.Fields("cantidad")

    End If

    If sww = 0 Then
        'If mytabley.State = 1 Then mytabley.Close
        mysql = ""
        mysql = "select * from almacen" & Chr$(10)
        mysql = mysql & "where local='" & Trim("" & mytablex.Fields("local")) & "'" & Chr$(10)
        mysql = mysql & "and producto='" & Trim("" & mytablex.Fields("producto")) & "'" & Chr$(10)
        mysql = mysql & "and bodega='" & Trim("" & mytablex.Fields("bodega")) & "'" & Chr$(10)
  
        mytabley.Open mysql, cn, adOpenStatic, adLockOptimistic
  
        'MsgBox mytabley.RecordCount
        If mytabley.RecordCount = 0 Then 'si existe
            'MsgBox ""
            mytabley.AddNew
            mytabley.Fields("local") = "" & mytablex.Fields("local")
            mytabley.Fields("producto") = "" & mytablex.Fields("producto")
            mytabley.Fields("bodega") = "" & mytablex.Fields("bodega")
            'sdx = Val("" & mytabley.Fields("saldo")) + signo * Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
            sdx = Val("" & mytabley.Fields("saldo")) + Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
            'MsgBox sdx
            mytabley.Fields("saldo") = sdx
            mytabley.Update
        Else

            If sw = 0 Then
                'sdx = Val("" & mytabley.Fields("saldo")) + signo * Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
                sdx = Val("" & mytabley.Fields("saldo")) + Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
                'MsgBox sdx
                mytabley.Fields("saldo") = sdx
                decarga_saldo_talla mytabley, mytablex, signo
                mytabley.Update

            End If

            If sw = 1 Then 'si es la primera vez pll
                sdx = Val("" & mytabley.Fields("saldo")) - Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
                mytabley.Fields("saldo") = sdx
                decarga_saldo_talla mytabley, mytablex, signo
                mytabley.Update

            End If

            'inicio 15/09/2017 pll
            If sw = 2 Then 'si es la primera vez pll
                If my_nuevaCantidad = sdx Then
                    mytabley.Fields("saldo") = sdx
                    mytabley.Update

                End If

            End If

            'fin 15/09/2017 pll
            '-------------------------------
        End If

    End If 'fin sw sw

    '-------------------------------------------------
    'inicio 15/09/2017 pll
    If sww = 1 Then
        If sw = 2 Then 'si es la primera vez pll

            'If my_nuevaCantidad = sdx Then
            'sdx = Val(my_nuevaCantidad)
            '  mytabley.Fields("saldo") = sdx
            '  mytabley.Update
            'End If
        End If

    End If

    'fin 15/09/2017 pll
    mytabley.Close
    mytablex.MoveNext
    'Loop
    Exit Sub
cmd19_err:
    MsgBox "Aviso en descarga saldo " + error$, 48, "Aviso"
    Exit Sub

End Sub

Public Sub decarga_saldo_talla(mytablex As ADODB.Recordset, _
                               mytabley As ADODB.Recordset, _
                               signo As Double)

    Dim sdx As Double

    sdx = Val("" & mytablex.Fields("t1")) + signo * Val("" & mytabley.Fields("t1"))
    mytablex.Fields("t1") = sdx
    sdx = Val("" & mytablex.Fields("t2")) + signo * Val("" & mytabley.Fields("t2"))
    mytablex.Fields("t2") = sdx
    sdx = Val("" & mytablex.Fields("t3")) + signo * Val("" & mytabley.Fields("t3"))
    mytablex.Fields("t3") = sdx
    sdx = Val("" & mytablex.Fields("t4")) + signo * Val("" & mytabley.Fields("t4"))
    mytablex.Fields("t4") = sdx
    sdx = Val("" & mytablex.Fields("t5")) + signo * Val("" & mytabley.Fields("t5"))
    mytablex.Fields("t5") = sdx
    sdx = Val("" & mytablex.Fields("t6")) + signo * Val("" & mytabley.Fields("t6"))
    mytablex.Fields("t6") = sdx
    sdx = Val("" & mytablex.Fields("t7")) + signo * Val("" & mytabley.Fields("t7"))
    mytablex.Fields("t7") = sdx
    sdx = Val("" & mytablex.Fields("t8")) + signo * Val("" & mytabley.Fields("t8"))
    mytablex.Fields("t8") = sdx
    sdx = Val("" & mytablex.Fields("t9")) + signo * Val("" & mytabley.Fields("t9"))
    mytablex.Fields("t9") = sdx
    sdx = Val("" & mytablex.Fields("t10")) + signo * Val("" & mytabley.Fields("t10"))
    mytablex.Fields("t10") = sdx
    sdx = Val("" & mytablex.Fields("t11")) + signo * Val("" & mytabley.Fields("t11"))
    mytablex.Fields("t11") = sdx
    sdx = Val("" & mytablex.Fields("t12")) + signo * Val("" & mytabley.Fields("t12"))
    mytablex.Fields("t12") = sdx
    sdx = Val("" & mytablex.Fields("t13")) + signo * Val("" & mytabley.Fields("t13"))
    mytablex.Fields("t13") = sdx
    sdx = Val("" & mytablex.Fields("t14")) + signo * Val("" & mytabley.Fields("t14"))
    mytablex.Fields("t14") = sdx
    sdx = Val("" & mytablex.Fields("t15")) + signo * Val("" & mytabley.Fields("t15"))
    mytablex.Fields("t15") = sdx
    sdx = Val("" & mytablex.Fields("t16")) + signo * Val("" & mytabley.Fields("t16"))
    mytablex.Fields("t16") = sdx

End Sub

'inicio 12/09/2017 pll
Public Sub carga_saldo(xlocal As String, _
                       xtipo As String, _
                       xserie As String, _
                       xnumero As String, _
                       sw As Integer, _
                       saldo_anterior, _
                       xtipo1 As String, _
                       acu)

    Dim sdx          As Double

    Dim signo        As Double

    Dim sww          As Integer

    Dim mytablestock As New ADODB.Recordset

    Dim mytablex     As New ADODB.Recordset

    Dim mytabley     As New ADODB.Recordset

    Dim buf          As String

    Dim found        As Integer

    Dim mysql        As String

    Dim my_saldoini  As Double

    On Error GoTo cmd19_err

    sww = 0
    'AHORA HAY QUE VALIDAR QUE no existe ya cruzado el documento----
  
    mysql = "SELECT * FROM " & cgusuario & Chr$(10)
    mysql = mysql & "where  local='" & xlocal & "' " & Chr$(10)
    mysql = mysql & "and tipo='" & xtipo & "'" & Chr$(10)
    mysql = mysql & "and serie='" & xserie & "'" & Chr$(10)
    mysql = mysql & "and numero='" & xnumero & "'" & Chr$(10)
   
    mytablestock.Open mysql, cn, adOpenStatic, adLockOptimistic
 
    If mytablestock.RecordCount > 0 Then  'si existe
        sww = 0

    End If

    'inicip 14/09/2017 pll
    'If acu = "C" And my_acu = "" Then
    '  buf = "dordenc"
    'ElseIf acu = "V" Or acu = "T" Or acu = "S" Then
    If acu = "V" Or acu = "T" Or acu = "S" Or acu = "Z" Or acu = "E" Or acu = "F" Then
        'If acu = "V" Or acu = "T" Or acu = "S" Then
        buf = "detalle"
    ElseIf acu = "C" Or my_acu = "S" Or my_acu = "K" Then
        buf = "detalle"

    End If

    mysql = ""
    mysql = "SELECT * FROM " & buf & Chr$(10)
    mysql = mysql & "where  local='" & xlocal & "'" & Chr$(10)
 
    If acu = "V" Or acu = "T" Or acu = "S" Or my_acu = "K" Or acu = "C" Or acu = "Z" Or acu = "E" Or acu = "F" Then
        mysql = mysql & "and tipo='" & xtipo & "' " & Chr$(10)

    End If

    If acu = "V" Or acu = "T" Or acu = "S" Or my_acu = "K" Or acu = "C" Or acu = "Z" Or acu = "E" Or acu = "F" Then
        mysql = mysql & "and serie='" & xserie & "'" & Chr$(10)

    End If

    If acu = "V" Or acu = "T" Or acu = "S" Or my_acu = "K" Or acu = "C" Or acu = "Z" Or acu = "E" Or acu = "F" Then
        mysql = mysql & "and numero='" & xnumero & "'" & Chr$(10)

    End If
 
    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic
  
    If mytablex.RecordCount = 0 Then 'si existe
        mytablex.Close
        Exit Sub

    End If

    Do

        If mytablex.EOF Then Exit Do
        If sww = 0 Then
            mysql = ""
            mysql = "select * from almacen" & Chr$(10)
            mysql = mysql & "where local='" & Trim("" & mytablex.Fields("local")) & "'" & Chr$(10)
            mysql = mysql & "and producto='" & Trim("" & mytablex.Fields("producto")) & "'" & Chr$(10)
            mysql = mysql & "and bodega='" & Trim("" & mytablex.Fields("bodega")) & "'" & Chr$(10)
  
            mytabley.Open mysql, cn, adOpenStatic, adLockOptimistic
  
            If mytabley.RecordCount = 0 Then 'si existe
                mytabley.AddNew
                mytabley.Fields("local") = "" & mytablex.Fields("local")
                mytabley.Fields("producto") = "" & mytablex.Fields("producto")
                mytabley.Fields("bodega") = "" & mytablex.Fields("bodega")
                sdx = Val("" & mytabley.Fields("saldo")) + Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
                mytabley.Fields("saldo") = sdx
                mytabley.Fields("entrada") = sdx
                mytabley.Fields("saldoinicial") = sdx
                mytabley.Update
                'inicio 19/09/2017 pll
                Call graba_producto(sdx, mytablex.Fields("producto"))
                'fin 19/09/2017 pll
                'Exit Sub
            Else
                '-------------------------------aqui control del producto
                Call saldo_producto(my_saldoini, mytablex.Fields("producto"))

                If sw = 1 And acu = "T" Then

                    'inicio 14/09/2017 pll
                    If mytabley.Fields("saldo") < mytablex.Fields("cantidad") Then
                        mytabley.Fields("saldo") = mytablex.Fields("cantidad")
                    Else
                        sdx = Val("" & mytabley.Fields("saldo")) + Val("" & mytabley.Fields("svirtual"))
                        mytabley.Fields("saldo") = sdx

                    End If

                    mytabley.Update

                End If

                'inicio 18/09/20117 pll para la factura de compras
                If sw = 1 And acu = "C" And xtipo1 = "Nuevo" Then
                    sdx = mytabley.Fields("saldo")
                    mytabley.Fields("saldo") = sdx + mytablex.Fields("cantidad")
                    mytabley.Fields("entrada") = mytablex.Fields("cantidad")
                    mytabley.Update

                End If

                If sw = 1 And acu = "C" And xtipo1 = "Modifica" Then
                    If mytabley.Fields("saldo") = mytablex.Fields("cantidad") And mytablex.Fields("cantidad") = mytabley.Fields("entrada") Then 'aqui paso modifica si son iguales las cantidades
                        If mytabley.Fields("saldo") = my_saldoini And mytabley.Fields("saldo") = mytabley.Fields("entrada") Then
                            mytabley.Fields("saldo") = mytablex.Fields("cantidad")
                            mytabley.Fields("entrada") = mytablex.Fields("cantidad")
                            mytabley.Fields("saldoinicial") = mytablex.Fields("cantidad")
                            mytabley.Update

                        End If

                        'Else
                        'MsgBox "xx"
                        'End If
                    ElseIf sw = 1 And acu = "C" And mytabley.Fields("saldo") <> mytablex.Fields("cantidad") And mytablex.Fields("cantidad") <> mytabley.Fields("entrada") And xtipo1 = "Modifica" Then

                        If sw = 1 And acu = "C" And mytabley.Fields("saldo") = my_saldoini And mytabley.Fields("saldo") = mytabley.Fields("entrada") Then
                            mytabley.Fields("saldo") = mytablex.Fields("cantidad")
                            mytabley.Fields("entrada") = mytablex.Fields("cantidad")
                            mytabley.Fields("saldoinicial") = mytablex.Fields("cantidad")
                            mytabley.Update
                        Else

                            If sw = 1 And acu = "C" And mytabley.Fields("entrada") < mytablex.Fields("cantidad") And xtipo1 = "Modifica" Then
                                sdx = mytabley.Fields("saldo") - mytabley.Fields("entrada")
                                mytabley.Fields("saldo") = sdx + mytablex.Fields("cantidad")
                                mytabley.Fields("entrada") = mytablex.Fields("cantidad")
                                mytabley.Update
                            Else
                                sdx = mytabley.Fields("saldo") - mytabley.Fields("entrada")
                                mytabley.Fields("saldo") = sdx + mytablex.Fields("cantidad")
                                mytabley.Fields("entrada") = mytablex.Fields("cantidad")
                                mytabley.Update

                            End If

                        End If

                    End If

                End If

                'control en caso que son iguales las cantidades con el saldo anterior
                If sw = 1 And acu = "C" And mytablex.Fields("cantidad") = saldo_anterior And xtipo1 = "Modifica" Then
                    Exit Sub
                End If 'cierre el control de las cantidades iguales

                'para anular y eliminar guia de salida
                If sw = 1 And acu = "C" And xtipo1 = "Eliminar" Then
                    ' new1 = mytabley.Fields("saldo") - saldo_anterior
                    new1 = mytabley.Fields("saldo") - my_saldoini
                    mytabley.Fields("saldo") = new1
                    mytabley.Update

                End If

            End If 'por mientras
     
        End If

        'Inicio 10/11/2017 pll para el pedido
        If sw = 1 And acu = "E" And xtipo1 = "Nuevo" Then
            sdx = mytabley.Fields("saldo")
            mytabley.Fields("saldo") = sdx + mytablex.Fields("cantidad")
            mytabley.Fields("entrada") = mytablex.Fields("cantidad")
            mytabley.Update

        End If

        If sw = 1 And acu = "E" And xtipo1 = "Modifica" Then
            If mytabley.Fields("saldo") = mytablex.Fields("cantidad") And mytablex.Fields("cantidad") = mytabley.Fields("entrada") Then 'aqui paso modifica si son iguales las cantidades
                If mytabley.Fields("saldo") = my_saldoini And mytabley.Fields("saldo") = mytabley.Fields("entrada") Then
                    mytabley.Fields("saldo") = mytablex.Fields("cantidad")
                    mytabley.Fields("entrada") = mytablex.Fields("cantidad")
                    mytabley.Fields("saldoinicial") = mytablex.Fields("cantidad")
                    mytabley.Update

                End If

            ElseIf sw = 1 And acu = "E" And mytabley.Fields("saldo") <> mytablex.Fields("cantidad") And mytablex.Fields("cantidad") <> mytabley.Fields("entrada") And xtipo1 = "Modifica" Then

                If sw = 1 And acu = "E" And mytabley.Fields("saldo") = my_saldoini And mytabley.Fields("saldo") = mytabley.Fields("entrada") Then
                    mytabley.Fields("saldo") = mytablex.Fields("cantidad")
                    mytabley.Fields("entrada") = mytablex.Fields("cantidad")
                    mytabley.Fields("saldoinicial") = mytablex.Fields("cantidad")
                    mytabley.Update
                Else

                    If sw = 1 And acu = "E" And mytabley.Fields("entrada") < mytablex.Fields("cantidad") And xtipo1 = "Modifica" Then
                        sdx = mytabley.Fields("saldo") - mytabley.Fields("entrada")
                        mytabley.Fields("saldo") = sdx + mytablex.Fields("cantidad")
                        mytabley.Fields("entrada") = mytablex.Fields("cantidad")
                        mytabley.Update
                    Else
                        sdx = mytabley.Fields("saldo") - mytabley.Fields("entrada")
                        mytabley.Fields("saldo") = sdx + mytablex.Fields("cantidad")
                        mytabley.Fields("entrada") = mytablex.Fields("cantidad")
                        mytabley.Update

                    End If

                End If

            End If

        End If

        'control en caso que son iguales las cantidades con el saldo anterior
        If sw = 1 And acu = "E" And mytablex.Fields("cantidad") = saldo_anterior And xtipo1 = "Modifica" Then
            Exit Sub
        End If 'cierre el control de las cantidades iguales

        'para anular y eliminar guia de salida
        If sw = 1 And acu = "E" And xtipo1 = "Eliminar" Then
            ' new1 = mytabley.Fields("saldo") - saldo_anterior
            new1 = mytabley.Fields("saldo") - my_saldoini
            mytabley.Fields("saldo") = new1
            mytabley.Update

        End If

        'inicio 07/12/20917 pll para la nota debito ventas
        If sw = 1 And acu = "F" And xtipo1 = "Nuevo" Then
            sdx = mytabley.Fields("saldo")
            mytabley.Fields("saldo") = sdx + mytablex.Fields("cantidad")
            mytabley.Fields("entrada") = mytablex.Fields("cantidad")
            mytabley.Update

        End If

        If sw = 1 And acu = "F" And xtipo1 = "Modifica" Then
            If mytabley.Fields("saldo") = mytablex.Fields("cantidad") And mytablex.Fields("cantidad") = mytabley.Fields("entrada") Then 'aqui paso modifica si son iguales las cantidades
                If mytabley.Fields("saldo") = my_saldoini And mytabley.Fields("saldo") = mytabley.Fields("entrada") Then
                    mytabley.Fields("saldo") = mytablex.Fields("cantidad")
                    mytabley.Fields("entrada") = mytablex.Fields("cantidad")
                    mytabley.Fields("saldoinicial") = mytablex.Fields("cantidad")
                    mytabley.Update

                End If

            ElseIf sw = 1 And acu = "E" And mytabley.Fields("saldo") <> mytablex.Fields("cantidad") And mytablex.Fields("cantidad") <> mytabley.Fields("entrada") And xtipo1 = "Modifica" Then

                If sw = 1 And acu = "F" And mytabley.Fields("saldo") = my_saldoini And mytabley.Fields("saldo") = mytabley.Fields("entrada") Then
                    mytabley.Fields("saldo") = mytablex.Fields("cantidad")
                    mytabley.Fields("entrada") = mytablex.Fields("cantidad")
                    mytabley.Fields("saldoinicial") = mytablex.Fields("cantidad")
                    mytabley.Update
                Else

                    If sw = 1 And acu = "F" And mytabley.Fields("entrada") < mytablex.Fields("cantidad") And xtipo1 = "Modifica" Then
                        sdx = mytabley.Fields("saldo") - mytabley.Fields("entrada")
                        mytabley.Fields("saldo") = sdx + mytablex.Fields("cantidad")
                        mytabley.Fields("entrada") = mytablex.Fields("cantidad")
                        mytabley.Update
                    Else
                        sdx = mytabley.Fields("saldo") - mytabley.Fields("entrada")
                        mytabley.Fields("saldo") = sdx + mytablex.Fields("cantidad")
                        mytabley.Fields("entrada") = mytablex.Fields("cantidad")
                        mytabley.Update

                    End If

                End If

            End If

        End If

        'control en caso que son iguales las cantidades con el saldo anterior
        If sw = 1 And acu = "F" And mytablex.Fields("cantidad") = saldo_anterior And xtipo1 = "Modifica" Then
            Exit Sub
        End If 'cierre el control de las cantidades iguales

        'para anular y eliminar guia de salida
        If sw = 1 And acu = "F" And xtipo1 = "Eliminar" Then
            new1 = mytabley.Fields("saldo") - my_saldoini
            mytabley.Fields("saldo") = new1
            mytabley.Update

        End If

        'fin 07/12/20917 pll
        'Fin 10/11/2017 pll
        '**fin 18/09/20117 pll para la guia salida
        'Call saldo_producto(my_saldoini, mytablex.Fields("producto"))
        If sw = 1 And acu = "S" And xtipo1 = "Nuevo" Then 'en caso de agregar una nueva guia salid
            sdx = mytabley.Fields("saldo")
            mytabley.Fields("saldo") = sdx + mytablex.Fields("cantidad")
            mytabley.Fields("entrada") = mytablex.Fields("cantidad")
            mytabley.Update

        End If

        If sw = 1 And acu = "S" And xtipo1 = "Modifica" Then 'guia entrada
            If mytabley.Fields("saldo") = mytablex.Fields("cantidad") And mytablex.Fields("cantidad") = mytabley.Fields("entrada") Then 'aqui paso modifica si son iguales las cantidades
                If mytabley.Fields("saldo") = my_saldoini And mytabley.Fields("saldo") = mytabley.Fields("salida") Then
                    mytabley.Fields("saldo") = mytablex.Fields("cantidad")
                    mytabley.Fields("entrada") = mytablex.Fields("cantidad")
                    mytabley.Fields("saldoinicial") = mytablex.Fields("cantidad")

                End If

            ElseIf sw = 1 And acu = "S" And xtipo1 = "Modifica" And mytabley.Fields("saldo") <> mytablex.Fields("cantidad") And mytablex.Fields("cantidad") <> mytabley.Fields("entrada") Then

            End If

            If sw = 1 And acu = "S" And xtipo1 = "Modifica" And mytabley.Fields("saldo") = my_saldoini And mytabley.Fields("saldo") = mytabley.Fields("entrada") Then
             
                mytabley.Fields("saldo") = mytablex.Fields("cantidad")
                mytabley.Fields("entrada") = mytablex.Fields("cantidad")
                mytabley.Fields("saldoinicial") = mytablex.Fields("cantidad")
            Else

                If sw = 1 And acu = "S" And mytabley.Fields("entrada") < mytablex.Fields("cantidad") And xtipo1 = "Modifica" Then
                    sdx = mytablex.Fields("cantidad") - mytabley.Fields("entrada")
                    new1 = mytabley.Fields("saldo") + sdx
                    mytabley.Fields("saldo") = new1
                    mytabley.Fields("entrada") = mytablex.Fields("cantidad")
                Else
                    sdx = mytabley.Fields("entrada") - mytablex.Fields("cantidad")
                    new1 = mytabley.Fields("saldo") + sdx
                    mytabley.Fields("saldo") = new1
                    mytabley.Fields("entrada") = mytablex.Fields("cantidad")

                End If

            End If

        End If

        'aqui el control de las cantidades con el saldo anterior
        If sw = 1 And acu = "S" And mytablex.Fields("cantidad") = saldo_anterior And xtipo1 = "Modifica" Then
            Exit Sub
        End If 'cierra el control de las cantidades

        If sw = 1 And acu = "S" And xtipo1 = "Eliminar" Then
            new1 = mytabley.Fields("saldo") - my_saldoini
            mytabley.Fields("saldo") = new1
            mytabley.Update

        End If

        mytabley.Update
        Call graba_producto(mytablex.Fields("cantidad"), mytablex.Fields("producto"))
        'End If 'fin sw sw
        '-------------------------------------------------
        mytabley.Close
        mytablex.MoveNext
    Loop
    mytablex.Close
    Exit Sub
cmd19_err:
    MsgBox "Aviso en descarga saldo " + error$, 48, "Aviso"
    Exit Sub

End Sub

'fin 12/0972017 pll
'inicio 16/09/2017 pll
Public Sub sacar_saldo(xlocal As String, _
                       xtipo As String, _
                       xserie As String, _
                       xnumero As String, _
                       sw As Integer, _
                       xtipo1 As String, _
                       acu)

    Dim sdx          As Double

    Dim signo        As Double

    Dim sww          As Integer

    Dim mytablestock As New ADODB.Recordset

    Dim mytablex     As New ADODB.Recordset

    Dim mytabley     As New ADODB.Recordset

    Dim mytables     As New ADODB.Recordset

    Dim buf          As String

    Dim found        As Integer

    Dim mysql        As String

    Dim my_saldoini  As Double

    Dim new1         As Double

    'Dim k                                                           As Integer

    On Error GoTo cmd19_err

    sww = 0
    'AHORA HAY QUE VALIDAR QUE no existe ya cruzado el documento----
  
    mysql = "SELECT * FROM " & cgusuario & Chr$(10)
    mysql = mysql & "where  local='" & xlocal & "' " & Chr$(10)
    mysql = mysql & "and tipo='" & xtipo & "'" & Chr$(10)
    mysql = mysql & "and serie='" & xserie & "'" & Chr$(10)
    mysql = mysql & "and numero='" & xnumero & "'" & Chr$(10)
   
    mytablestock.Open mysql, cn, adOpenStatic, adLockOptimistic
 
    If mytablestock.RecordCount > 0 Then  'si existe
        sww = 0

    End If

    'inicip 14/09/2017 pll
    If acu = "C" Then
        buf = "dordenc"
    ElseIf acu = "V" Or acu = "T" Then
        buf = "detalle"
    ElseIf acu = "Z" Then
        'buf = "DTRASLAD"
        buf = "detalle"

    End If

    mysql = ""
    mysql = "SELECT * FROM " & buf & Chr$(10)
    mysql = mysql & "where  local='" & xlocal & "'" & Chr$(10)

    If acu = "V" Then
        mysql = mysql & "and tipo='" & xtipo & "' " & Chr$(10)
    ElseIf acu = "C" Then
        mysql = mysql & "and tipo='OC' " & Chr$(10)
    ElseIf acu = "T" Then
        mysql = mysql & "and tipo='T' " & Chr$(10)
    ElseIf acu = "Z" Then
        mysql = mysql & "and tipo='" & xtipo & "'" & Chr$(10)

    End If
 
    If acu = "V" Or acu = "T" Or acu = "Z" Then
        mysql = mysql & "and serie='" & xserie & "'" & Chr$(10)
    ElseIf acu = "C" Then
        mysql = mysql & "and serie='" & my_serie1 & "'" & Chr$(10)

    End If
 
    If acu = "V" Or acu = "T" Or acu = "Z" Then
        mysql = mysql & "and numero='" & xnumero & "'" & Chr$(10)
    ElseIf acu = "C" Then
        mysql = mysql & "and numero='" & my_numero1 & "'" & Chr$(10)

    End If
 
    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic
  
    If mytablex.RecordCount = 0 Then 'si existe
        Exit Sub

    End If

    Do

        If mytablex.EOF Then Exit Do
        If sww = 0 Then
            mysql = ""
            mysql = "select * from almacen" & Chr$(10)
            mysql = mysql & "where local='" & Trim("" & mytablex.Fields("local")) & "'" & Chr$(10)
            mysql = mysql & "and producto='" & Trim("" & mytablex.Fields("producto")) & "'" & Chr$(10)
            mysql = mysql & "and bodega='" & Trim("" & mytablex.Fields("bodega")) & "'" & Chr$(10)
  
            mytabley.Open mysql, cn, adOpenStatic, adLockOptimistic
  
            If mytabley.RecordCount = 0 Then 'si existe
                mytabley.AddNew
                mytabley.Fields("local") = "" & mytablex.Fields("local")
                mytabley.Fields("producto") = "" & mytablex.Fields("producto")
                mytabley.Fields("bodega") = "" & mytablex.Fields("bodega")

                If acu = "C" Then
                    sdx = Val("" & mytabley.Fields("saldo")) + Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
                ElseIf acu = "V" Then
                    Call saldo_producto(my_saldoini, mytablex.Fields("producto"))

                    If my_saldoini = 0 Then
                        sdx = Val("" & mytabley.Fields("saldo")) - Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))

                    End If

                End If

                mytabley.Fields("saldo") = sdx
                mytabley.Fields("entrada") = sdx
                mytabley.Fields("saldoinicial") = sdx
                my_saldoini = sdx
                mytabley.Update
                sw = 0
                'Exit Sub
            Else
     
                '-------------------------------aqui control del producto
                'mytablex("cantidad") --> detalle
                'mytabley("salida") -->almacen
                Call saldo_producto(my_saldoini, mytablex.Fields("producto"))

                '**guia de salida
                If sw = 1 And acu = "T" And xtipo1 = "Nuevo" Then   ' si es la primera vez pll
                    sdx = mytabley.Fields("saldo")
                    mytabley.Fields("saldo") = sdx - mytablex.Fields("cantidad")
                    mytabley.Fields("salida") = Val("" & mytablex.Fields("cantidad"))
                    mytabley.Fields("svirtual") = mytablex("cantidad")

                End If

                If sw = 1 And acu = "T" And xtipo1 = "Modifica" Then
                    If mytabley.Fields("saldo") = mytablex.Fields("cantidad") And mytablex.Fields("cantidad") = mytabley.Fields("salida") Then
                        If mytabley.Fields("saldo") = my_saldoini And mytabley.Fields("saldo") = mytabley.Fields("salida") Then
                            mytabley.Fields("saldo") = mytablex.Fields("cantidad")
                            mytabley.Fields("salida") = mytablex.Fields("cantidad")
                            mytabley.Fields("svirtual") = mytabley.Fields("salida")

                        End If

                    ElseIf sw = 1 And acu = "T" And xtipo1 = "Modifica" And mytabley.Fields("saldo") <> mytablex.Fields("cantidad") And mytablex.Fields("cantidad") <> mytabley.Fields("salida") Then

                        If my_saldoini > 0 Then
                            If mytabley.Fields("saldo") = my_saldoini And mytabley.Fields("saldo") = mytabley.Fields("salida") Then
                                mytabley.Fields("saldo") = mytablex.Fields("cantidad")
                                mytabley.Fields("entrada") = mytablex.Fields("cantidad")
                                mytabley.Fields("svirtual") = mytablex("cantidad")
                            Else 'encaso salida es mayor que se guardado

                                If sw = 1 And acu = "T" And xtipo1 = "Modifica" And mytablex.Fields("cantidad") < saldo_anterior Then
                                    sdx = mytablex.Fields("cantidad") - mytabley.Fields("salida")
                                    new1 = mytabley.Fields("saldo") - sdx
                                    mytabley.Fields("saldo") = new1
                                    mytabley.Fields("svirtual") = mytablex.Fields("cantidad")
                                    mytabley.Fields("salida") = mytablex.Fields("cantidad")
                                Else 'aqui para la salida que es menor
                                    sdx = mytabley.Fields("salida") - mytablex.Fields("cantidad")
                                    new1 = mytabley.Fields("saldo") + sdx
                                    mytabley.Fields("saldo") = new1
                                    mytabley.Fields("svirtual") = mytablex.Fields("cantidad")
                                    mytabley.Fields("salida") = mytablex.Fields("cantidad")

                                End If

                            End If

                        End If 'es del saldo

                    End If

                End If

                If sw = 1 And acu = "T" And mytablex.Fields("cantidad") = saldo_anterior And xtipo1 = "Modifica" Then
                    Exit Sub

                End If

                'para anular y eliminar guia de salida
                If sw = 1 And acu = "T" And xtipo1 = "Eliminar" Then
                    new1 = mytabley.Fields("saldo") + my_saldoini
                    mytabley.Fields("saldo") = new1

                End If
     
            End If ' cierra si son iguales cantidades con el saldo anterior

            '-------------------------------
            ' End If
            ''para las ventas
            If sw = 1 And acu = "V" And xtipo1 = "Nuevo" Then
                sdx = mytabley.Fields("saldo")
                mytabley.Fields("saldo") = sdx - mytablex.Fields("cantidad")
                mytabley.Fields("salida") = mytablex.Fields("cantidad")
                mytabley.Fields("svirtual") = mytablex("cantidad")
                my_saldoini = mytabley.Fields("saldo")

            End If

            If sw = 2 And acu = "V" And xtipo1 = "Nuevo" Then
                sdx = mytabley.Fields("saldo")
                mytabley.Fields("saldo") = sdx - mytablex.Fields("cantidad")
                mytabley.Fields("salida") = mytablex.Fields("cantidad")
                mytabley.Fields("svirtual") = mytablex("cantidad")
                my_saldoini = mytabley.Fields("saldo")
                sdx = mytabley.Fields("saldo")

            End If
     
            If sw = 1 And acu = "V" And xtipo1 = "Modifica" Then
                If mytabley.Fields("saldo") = mytablex.Fields("cantidad") And mytablex.Fields("cantidad") = mytabley.Fields("salida") Then 'aqui paso modifica si son iguales las cantidades
                    If mytabley.Fields("saldo") = my_saldoini And mytabley.Fields("saldo") = mytabley.Fields("salida") Then
                        mytabley.Fields("saldo") = mytablex.Fields("cantidad")
                        mytabley.Fields("salida") = mytablex.Fields("cantidad")
                        mytabley.Fields("svirtual") = mytabley.Fields("salida")

                    End If

                ElseIf sw = 1 And acu = "V" And xtipo1 = "Modifica" And mytabley.Fields("saldo") <> mytablex.Fields("cantidad") And mytablex.Fields("cantidad") <> mytabley.Fields("salida") Then

                    If mytabley.Fields("saldo") = my_saldoini And mytabley.Fields("saldo") = mytabley.Fields("salida") Then
                        mytabley.Fields("saldo") = mytablex.Fields("cantidad")
                        mytabley.Fields("entrada") = mytablex.Fields("cantidad")
                        mytabley.Fields("svirtual") = mytablex("cantidad")
                    Else 'encaso salida es mayor que se guardado

                        If my_saldoini > 0 Then
                            If sw = 1 And acu = "V" And mytabley.Fields("salida") < mytablex.Fields("cantidad") And xtipo1 = "Modifica" Then
                                sdx = mytablex.Fields("cantidad") - mytabley.Fields("salida")
                                new1 = mytabley.Fields("saldo") - sdx
                                mytabley.Fields("saldo") = new1
                                mytabley.Fields("svirtual") = mytablex.Fields("cantidad")
                                mytabley.Fields("salida") = mytablex.Fields("cantidad")
                            Else 'aqui para la salida que es menor
                                sdx = mytabley.Fields("salida") - mytablex.Fields("cantidad")
                                new1 = mytabley.Fields("saldo") + sdx
                                mytabley.Fields("saldo") = new1

                                mytabley.Fields("svirtual") = mytablex.Fields("cantidad")
                                mytabley.Fields("salida") = mytablex.Fields("cantidad")

                            End If

                        End If

                    End If '

                End If

            End If

            'control de las contidades si son iguales
            If sw = 1 And acu = "V" And mytablex.Fields("cantidad") = saldo_anterior And xtipo1 = "Modifica" Then
                Exit Sub
            End If 'cierra ventas modifica
      
            If sw = 1 And acu = "V" And xtipo1 = "Eliminar" Then
                If mytabley.Fields("saldo") < 0 Then
                    new1 = mytabley.Fields("saldo")
                    signo = -1

                    If my_saldoini = new1 Then
                        sdx = 0 + signo * Val(new1) * 0
                        mytabley.Fields("saldo") = sdx
                        my_saldoini = sdx
                    Else
                        sdx = 0 + signo * Val(new1) ' * my_saldoini
                        mytabley.Fields("saldo") = sdx

                    End If

                    'sdx = Val("" & mytabley.Fields("saldo")) + signo * Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
                Else
                    new1 = mytabley.Fields("saldo") + my_saldoini
                    mytabley.Fields("saldo") = new1

                End If

            End If
    
            '*aqui para la cotizacion
            If sw = 1 And acu = "Z" And xtipo = "TS" And xtipo1 = "Nuevo" Then   ' si es la primera vez pll
         
                sdx = mytabley.Fields("saldo")
                mytabley.Fields("saldo") = sdx - mytablex.Fields("cantidad")
                mytabley.Fields("salida") = Val("" & mytablex.Fields("cantidad"))
                mytabley.Fields("svirtual") = mytablex("cantidad")
                mytabley.Update

            End If

            If sw = 1 And acu = "Z" And xtipo = "TE" And xtipo1 = "Nuevo" Then   ' si es la primera vez pll
                sdx = mytabley.Fields("saldo")
                mytabley.Fields("saldo") = sdx + mytablex.Fields("cantidad")
                mytabley.Fields("salida") = Val("" & mytablex.Fields("cantidad"))
                mytabley.Fields("svirtual") = mytablex("cantidad")
                mytabley.Update

            End If

            'inicio 31/01/2018 pll en el caso que ingresa bodega secundaria
            If sw = 0 And acu = "Z" And xtipo = "TE" And xtipo1 = "Nuevo" Then
                'ElseIf sw = 1 And acu = "Z" And xtipo = "TE" And xtipo1 = "Modifica" And my_saldoini < mytablex.Fields("cantidad") Then
                sdx = mytablex.Fields("cantidad") - mytabley.Fields("saldo")
                new1 = mytabley.Fields("saldo") + sdx
                mytabley.Fields("saldo") = new1
                mytabley.Fields("entrada") = mytablex.Fields("cantidad")
                mytabley.Update

                'mytabley.Fields("salida") = mytablex.Fields("cantidad")
            End If

            'fin 31/01/2018 pll
            If sw = 1 And acu = "Z" And xtipo1 = "Modifica" Then
                If mytabley.Fields("saldo") = mytablex.Fields("cantidad") And mytablex.Fields("cantidad") = mytabley.Fields("salida") Then
                    If mytabley.Fields("saldo") = my_saldoini And mytabley.Fields("saldo") = mytabley.Fields("salida") Then
                        mytabley.Fields("saldo") = mytablex.Fields("cantidad")
                        mytabley.Fields("salida") = mytablex.Fields("cantidad")
                        mytabley.Fields("svirtual") = mytabley.Fields("salida")
                        mytabley.Update

                    End If

                ElseIf sw = 1 And acu = "Z" And xtipo1 = "Modifica" And mytabley.Fields("saldo") <> mytablex.Fields("cantidad") And mytablex.Fields("cantidad") <> mytabley.Fields("salida") Then

                    If mytabley.Fields("saldo") = my_saldoini And mytabley.Fields("saldo") = mytabley.Fields("salida") Then
                        mytabley.Fields("saldo") = mytablex.Fields("cantidad")
                        mytabley.Fields("entrada") = mytablex.Fields("cantidad")
                        mytabley.Fields("svirtual") = mytablex("cantidad")
                        mytabley.Update
                    Else 'encaso salida es mayor que se guardado

                        'If sw = 1 And acu = "Z" And xtipo = "TS" And xtipo1 = "Modifica" And mytabley.Fields("salida") < mytablex.Fields("cantidad") Then
                        If sw = 1 And acu = "Z" And xtipo = "TS" And xtipo1 = "Modifica" And my_saldoini <> mytablex.Fields("cantidad") Then
                            sdx = mytablex.Fields("cantidad") - mytabley.Fields("salida")
                            new1 = mytabley.Fields("saldo") - sdx
                            mytabley.Fields("saldo") = new1
                            mytabley.Fields("svirtual") = mytablex.Fields("cantidad")
                            mytabley.Fields("salida") = mytablex.Fields("cantidad")
                            mytabley.Update

                        End If

                    End If

                ElseIf sw = 1 And acu = "Z" And xtipo = "TE" And xtipo1 = "Modifica" Then
                    'ElseIf sw = 1 And acu = "Z" And xtipo = "TE" And xtipo1 = "Modifica" And my_saldoini < mytablex.Fields("cantidad") Then
                    sdx = mytablex.Fields("cantidad") - mytabley.Fields("saldo")
                    new1 = mytabley.Fields("saldo") + sdx
                    mytabley.Fields("saldo") = new1
                    mytabley.Fields("entrada") = mytablex.Fields("cantidad")
                    mytabley.Update

                    'mytabley.Fields("salida") = mytablex.Fields("cantidad")
                End If

            End If

            'Translado es para anular
            If sw = 1 And acu = "Z" And xtipo = "TS" And xtipo1 = "Eliminar" Then
                new1 = mytabley.Fields("saldo") + my_saldoini
                'MsgBox "new1·" & new1
                mytabley.Fields("saldo") = new1
                mytabley.Update

            End If

            If sw = 1 And acu = "Z" And xtipo = "TE" And xtipo1 = "Eliminar" Then
                new1 = mytabley.Fields("saldo") - my_saldoini
                mytabley.Fields("saldo") = new1

            End If

            '-------------------------------
            mytabley.Update

            If sw = 1 And acu = "Z" And xtipo = "TS" And xtipo1 = "Nuevo" Then
                'inicio 17/01/2018 pll
                mytabley.Update
                Call graba_producto(mytablex.Fields("cantidad"), mytablex.Fields("producto"))
                'fin 17/01/2018 pll
            ElseIf my_saldoini < 0 Or my_saldoini = 0 Then
                mytabley.Update
                'mytablex.Fields("cantidad") = my_saldoini
                Call graba_producto(my_saldoini, mytablex.Fields("producto"))
                sw = 2
            Else
                mytabley.Update
                Call graba_producto(mytablex.Fields("cantidad"), mytablex.Fields("producto"))
    
                'End If
            End If

        End If 'fin sw sw

        '-------------------------------------------------
        mytabley.Close
        mytablex.MoveNext
    Loop
    Exit Sub
cmd19_err:
    MsgBox "Aviso en descarga saldo " + error$, 48, "Aviso"
    Exit Sub

End Sub

'fin 16/09/2017 pll
Public Sub grabarAnular()

    Dim rs               As Recordset

    Dim I                As Integer

    Dim pracu            As String

    Dim buf1             As String

    Dim found            As Integer

    Dim mytablex         As New ADODB.Recordset

    Dim mytabley         As New ADODB.Recordset

    Dim mytablez         As New ADODB.Recordset

    Dim mytablea         As New ADODB.Recordset

    Dim mytableb         As New ADODB.Recordset

    Dim mytablexy        As New ADODB.Recordset

    Dim te               As String

    Dim ts               As String

    Dim xc1              As Double

    Dim xc2              As Double

    Dim xc3              As Double

    Dim xc4              As Double

    Dim fila             As Integer

    Dim sw               As Integer

    Dim xbuf             As String

    Dim mysql            As String

    Dim my_nuevaCantidad As Double

    On Error GoTo cmd761_err

    'graba cabecera
    If Not IsNumeric(Numero) Then
        serie.SetFocus
        Exit Sub

    End If

    sw = 0

    If racu = "Z" Then  'abrir base datos traslado
        mytableb.Open "SELECT * FROM detalle where local='" & localf & "' and tipo='TS' and serie='" & serie & "' and numero='" & Numero & "'", cn, adOpenKeyset, adLockOptimistic

    End If

    'MsgBox dgusuariog
    'aqui es factura
    xbuf = "SELECT * FROM " & cgusuario & " where  local='" & my_local & "' and tipo='" & my_tipo & "' and serie='" & my_serie & "' and numero='" & my_numero & "'"
   
    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open xbuf, cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then  'si existe
        grabar = 1
    Else
        grabar = 1

    End If

    mytablex.Close

    'GRABANDO EN detalle
    mysql = "SELECT * FROM " & dgusuariog & Chr$(10)
    mysql = mysql & "where local='" & my_local & "'" & Chr$(10)
    mysql = mysql & "and tipo='" & my_tipo & "'" & Chr$(10)
    mysql = mysql & " and serie='" & my_serie & "' " & Chr$(10)
    mysql = mysql & "and numero='" & my_numero & "'" & Chr$(10)
 
    mytablexy.Open mysql, cn, adOpenStatic, adLockOptimistic

    Do
        carga_saldo my_local, my_tipo, my_serie, my_numero, 1, "", "" & my_tipo1, acu
    
        '----
    Loop
    mytablexy.Close
    Exit Sub
cmd761_err:
    MsgBox "Aviso en grabar " + error$, 48, "Aviso"
    Exit Sub

End Sub

Public Sub graba_producto(my_saldoini As Double, my_producto As String)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    mysql = "update producto" & Chr$(10)
    mysql = mysql & "set saldoini =" & my_saldoini & Chr$(10)
    mysql = mysql & "where producto='" & my_producto & "'" & Chr$(10)
    cn.Execute (mysql)

End Sub

Public Sub saldo_producto(my_saldoini As Double, my_producto As String)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    mysql = "select saldoini" & producto & Chr$(10)
    mysql = mysql & "from producto" & Chr$(10)
    mysql = mysql & "where producto='" & my_producto & "'" & Chr$(10)

    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic
 
    If mytablex.EOF Then
        'salida = False
        Exit Sub
    Else

        If mytablex.Fields("saldoini") <> "" Then
            my_saldoini = mytablex.Fields("saldoini")
     
        Else
            my_saldoini = 0

        End If

    End If

End Sub

Public Sub actualiza_stock(local1 As String, _
                           ttipo As String, _
                           serie As String, _
                           Numero As String, _
                           my_struc_Stock_SAnterior() As struc_Stock_SAnterior, _
                           k As Integer)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    ReDim my_struc_Stock_SAnterior(0)
    'mysql = "SELECT NRO_ITEMS FROM " & cgusuario & "" & Chr$(10)
    mysql = "SELECT producto,cantidad FROM  detalle" & Chr$(10)
    mysql = mysql & "where  local='" & local1 & "'" & Chr$(10)
    mysql = mysql & "and tipo='" & ttipo & "'" & Chr$(10)
    mysql = mysql & "and serie='" & serie & "'" & Chr$(10)
    mysql = mysql & "and numero='" & Numero & "'" & Chr$(10)
   
    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open mysql, cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then  'si existe
    Else

        Do Until mytablex.EOF

            If k > 0 Then
                ReDim Preserve my_struc_Stock_SAnterior(UBound(my_struc_Stock_SAnterior) + 1)

            End If

            'saldo_anterior = mytablex.Fields("NRO_ITEMS")
            If mytablex.Fields("producto") <> "" Then
                my_struc_Stock_SAnterior(UBound(my_struc_Stock_SAnterior)).producto = mytablex.Fields("producto")
            Else
                my_struc_Stock_SAnterior(UBound(my_struc_Stock_SAnterior)).producto = ""

            End If

            If mytablex.Fields("cantidad") <> "" Then
                my_struc_Stock_SAnterior(UBound(my_struc_Stock_SAnterior)).cantidad = mytablex.Fields("cantidad")
            Else
                my_struc_Stock_SAnterior(UBound(my_struc_Stock_SAnterior)).cantidad = ""

            End If

            k = k + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

End Sub

Public Sub b_estado(local1 As String, _
                    ttipo As String, _
                    serie As String, _
                    Numero As String, _
                    my_estado As String)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    mysql = ""
    mysql = "select f.estado " & Chr$(10)
    mysql = mysql & "from " & cgusuario & " f" & Chr$(10)
    mysql = mysql & " where f.local ='" & extra_loquesea(local1) & "'" & Chr$(10)
    mysql = mysql & " and f.tipo = '" & extra_loquesea(ttipo) & "'" & Chr$(10)
    mysql = mysql & " and f.serie = '" & extra_loquesea(serie) & "'" & Chr$(10)
    mysql = mysql & " and f.numero= '" & extra_loquesea(Numero) & "'" & Chr$(10)
 
    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF Then
        Exit Sub
    Else
        my_estado = mytablex.Fields("estado")

    End If

    mytablex.Close

    Exit Sub
cmd921_err:
    MsgBox "aviso en sql_cabeza   " & error$, 48, "Aviso"
    Exit Sub

End Sub

