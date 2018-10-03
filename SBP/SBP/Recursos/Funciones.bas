Attribute VB_Name = "Funciones"

' 26/07/2018 Desactivar Facturacion Electronica
Function Obtiene_EstadoSistema() As String
    Obtiene_EstadoSistema = "FE BYH"

    On Error GoTo cmd8912_err

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT estadosistema FROM parame where codigo='01'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        Obtiene_EstadoSistema = "" & mytablex.Fields("estadosistema")

    End If

    mytablex.Close
cmd8912_err:

End Function

' 26/07/2018 Desactivar Facturacion Electronica

'''' 25/07/2018 Delivery y Para Llevar desde mozo
Function busca_tipoSalon(salon As String)

    Dim mytabley As New ADODB.Recordset

    mytabley.Open "SELECT  tipo FROM salon where salon='" & Trim("" & salon) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        busca_tipoSalon = "" & Trim$("" & mytabley.Fields("tipo"))

    End If

    mytabley.Close

End Function

'''' 25/07/2018 Delivery y Para Llevar desde mozo

'''' 25/07/2018 Delivery y Para Llevar desde mozo
Function busca_CLienteXMesa(salon As String, mesa As String)

    Dim mytabley As New ADODB.Recordset

    mytabley.Open "SELECT  codigo FROM mesa where salon='" & Trim("" & salon) & "'  and mesa='" & Trim("" & mesa) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        busca_CLienteXMesa = "" & Trim$("" & mytabley.Fields("codigo"))

    End If

    mytabley.Close

End Function

'''' 25/07/2018 Delivery y Para Llevar desde mozo

' 17/07/2018 Factura de Exportación
Function busca_OpcionExportacion()

    Dim mytabley As New ADODB.Recordset

    mytabley.Open "SELECT OpcionExportacion FROM parame where codigo='01'", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        busca_OpcionExportacion = "" & Trim$("" & mytabley.Fields("OpcionExportacion"))

    End If

    mytabley.Close

End Function

' 17/07/2018 Factura de Exportación

'' 10/07/2018 Edicion Comanda
Function busca_OpcionNombre()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT opcionnombre FROM parame where  codigo='01'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        If mytablex.Fields("opcionnombre") = "CO" Or mytablex.Fields("opcionnombre") = "DC" Then
            busca_OpcionNombre = "" & mytablex.Fields("opcionnombre")
        Else
            busca_OpcionNombre = "DL"

        End If
      
    End If

    mytablex.Close

End Function

'' 10/07/2018 Edicion Comanda

'' 10/07/2018 Edicion Comanda
Function busca_TamañoComanda()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT TamanoComanda FROM parame where  codigo='01'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        If mytablex.Fields("TamanoComanda") = "8" Then
            busca_TamañoComanda = "8"
        ElseIf mytablex.Fields("TamanoComanda") = "10" Then
            busca_TamañoComanda = "10"
        ElseIf mytablex.Fields("TamanoComanda") = "12" Then
            busca_TamañoComanda = "12"
        Else
            busca_TamañoComanda = "20"

        End If
      
    End If

    mytablex.Close

End Function

'' 10/07/2018 Edicion Comanda

'07/08/2018 No descuenta stock en guia de remision
Function busca_DescuentaStock(tipo As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT DescuentaStock FROM tipo where  tipo='" & Trim("" & tipo) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        busca_DescuentaStock = "" & mytablex.Fields("DescuentaStock")

    End If

    mytablex.Close

End Function

'07/08/2018 No descuenta stock en guia de remision

'13/08/2018 Integración FE - Pizzeria
'13/08/2018 Integración FE - Pizzeria
'11/06/2018 Actualiza Precio Promedio Ponderado Masivo
''' 11/12/2017 SubReceta
Function OpcionActualizaCostoReceta(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT costoreceta FROM producto where producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If "" & mytablex.Fields("costoreceta") = "S" Then
            OpcionActualizaCostoReceta = "S"

        End If

    End If

    mytablex.Close

End Function

Function OpcionTipoCostoReceta()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT tcostoreceta FROM parame where codigo='01'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
   
        If "" & mytablex.Fields("tcostoreceta") = "CP" Then
            OpcionTipoCostoReceta = "" & mytablex.Fields("tcostoreceta")
        Else
            OpcionTipoCostoReceta = "CU"

        End If

    End If

    mytablex.Close

End Function

Function actualiza_CostoTotalReceta(buf As String, bufc As String)

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim suma     As Double

    suma = 0
                  
    If mytablex.State = 1 Then mytablezx.Close
    mytablex.Open "SELECT * FROM receta where  LINEA='' and PRODUCTOI='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        Do

            If mytablex.EOF Then Exit Do
                          
            If mytabley.State = 1 Then mytabley.Close
            mytabley.Open "SELECT * FROM receta where  LINEA='' and PRODUCTO='" & mytablex.Fields("producto") & "'", cn, adOpenDynamic, adLockOptimistic
            suma = 0

            If mytabley.RecordCount > 0 Then
                Do

                    If mytabley.EOF Then Exit Do
                    suma = suma + Val("" & mytabley.Fields("cantidad")) * Val("" & mytabley.Fields("precio"))
                    mytabley.MoveNext
                Loop
                cn.Execute ("update producto set costou=" & Val(suma) & ", costop=" & Val(suma) & "  where producto='" & mytablex.Fields("producto") & "'")
                actualiza_CostoTotalReceta2 mytablex.Fields("producto")
                    
            End If

            mytabley.Close
                    
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
                
    Exit Function
cmd9093_err:
    Exit Function

End Function

Function actualiza_CostoTotalReceta2(buf As String)

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim suma     As Double

    suma = 0

    Dim costo As Double

    costo = 0

    If mytablex.State = 1 Then mytablezx.Close
    mytablex.Open "SELECT * FROM receta where  LINEA='' and PRODUCTOI='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        Do

            If mytablex.EOF Then Exit Do

            If mytabley.State = 1 Then mytabley.Close
            mytabley.Open "SELECT * FROM receta where  LINEA='' and PRODUCTO='" & mytablex.Fields("producto") & "'", cn, adOpenDynamic, adLockOptimistic
            suma = 0

            If mytabley.RecordCount > 0 Then
                Do

                    If mytabley.EOF Then Exit Do
                    'suma = suma + Val("" & mytabley.Fields("cantidad")) * Val("" & mytabley.Fields("precio"))
                    costo = Obtiene_CostoProducto(mytabley.Fields("productoi"))
                    cn.Execute ("update receta set precio=" & Val(costo) & ", total=cantidad* " & Val(costo) & "  where productoi='" & mytabley.Fields("productoi") & "' and producto='" & mytablex.Fields("producto") & "'")
 
                    cn.Execute ("update producto set costou=" & Val(suma) & ", costop=" & Val(suma) & "  where producto='" & mytablex.Fields("producto") & "'")
                    Actualiza_CostoProducto (mytablex.Fields("producto"))
                    Actualiza_CostoProducto2 (mytablex.Fields("producto"))
                  
                    mytabley.MoveNext
                Loop
                       
            End If

            mytabley.Close

            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

    Exit Function
cmd9093_err:
    Exit Function

End Function

Function Obtiene_CostoProducto(buf As String)

    Dim mytablex As New ADODB.Recordset
       
    If mytablex.State = 1 Then mytablezx.Close
    mytablex.Open "SELECT costop,costou FROM producto where PRODUCTO='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        Obtiene_CostoProducto = "" & mytablex.Fields("costop")

        'cu = "" & mytabley.Fields("costou")
    End If

    mytablex.Close
                
    Exit Function

End Function

Function Actualiza_CostoProducto(buf As String)

    Dim mytabley As New ADODB.Recordset

    Dim suma     As Double

    suma = 0
         
    If mytabley.State = 1 Then mytabley.Close
    mytabley.Open "SELECT * FROM receta where  LINEA='' and PRODUCTO='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
    suma = 0

    If mytabley.RecordCount > 0 Then
        Do

            If mytabley.EOF Then Exit Do
                           
            suma = suma + Val("" & mytabley.Fields("cantidad")) * Val("" & mytabley.Fields("precio"))
            mytabley.MoveNext
                            
        Loop
        cn.Execute ("update producto set costou=" & Val(suma) & ", costop=" & Val(suma) & "  where producto='" & buf & "'")
                            
    End If

    mytabley.Close
cmd9093_err:
    Exit Function

End Function

Function Actualiza_CostoProducto2(buf As String)

    Dim mytabley As New ADODB.Recordset

    Dim suma     As Double

    suma = 0

    Dim costo As Double

    costo = 0
         
    If mytabley.State = 1 Then mytabley.Close
    mytabley.Open "SELECT * FROM receta where  LINEA='' and PRODUCTOi='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
    suma = 0

    If mytabley.RecordCount > 0 Then
        Do

            If mytabley.EOF Then Exit Do
                           
            costo = Obtiene_CostoProducto(buf)
            cn.Execute ("update RECETA set PRECIO=" & Val(costo) & ", total=CANTIDAD * " & Val(costo) & "  where producto='" & Val("" & mytabley.Fields("producto")) & "' AND productoi='" & buf & "' ")
            Actualiza_CostoProducto Val("" & mytabley.Fields("producto"))
                        
            mytabley.MoveNext
               
        Loop

        'cn.Execute ("update producto set costou=" & Val(suma) & ", costop=" & Val(suma) & "  where producto='" & buf & "'")
    End If

    mytabley.Close
cmd9093_err:
    Exit Function

End Function

'13/08/2018 Integración FE - Pizzeria
'13/08/2018 Integración FE - Pizzeria

'15/08/2018 Cambiar Descripcion de producto venta de ventas
Function busca_CambiaDescripcionVentas(OPCION As Integer)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT Cambiadescripcion,nuevoproducto FROM parame where  codigo='01'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        If OPCION = 1 Then
            If mytablex.Fields("Cambiadescripcion") = "S" Then
                busca_CambiaDescripcionVentas = "S"
            Else
                busca_CambiaDescripcionVentas = "N"

            End If

        ElseIf OPCION = 2 Then

            If mytablex.Fields("nuevoproducto") = "S" Then
                busca_CambiaDescripcionVentas = "S"
            Else
                busca_CambiaDescripcionVentas = "N"

            End If

        End If
      
    End If

    mytablex.Close

End Function

'15/08/2018 Cambiar Descripcion de producto venta de ventas

'24/08/2018  Delivery por mesa
Function busca_EstadoXMesa(salon As String, mesa As String)

    Dim mytabley As New ADODB.Recordset

    mytabley.Open "SELECT  * FROM dcomanda where salon='" & Trim("" & salon) & "'  and mesa='" & Trim("" & mesa) & "'", cn, adOpenKeyset, adLockOptimistic
    busca_EstadoXMesa = "L" 'OCUPADA

    If mytabley.RecordCount > 0 Then
        busca_EstadoXMesa = "O" 'LIBRE

    End If

    mytabley.Close

End Function

'24/08/2018  Delivery por mesa

