Attribute VB_Name = "ElectronicoSunat"

'inicio 20/11/2017 pll
Type struc_ubigeo_Receptor

    idubigeo                        As String
    provincia                       As String
    departamento                    As String
    distrito                        As String
    direccion                       As String

End Type

Global my_struc_ubigeo_Receptor() As struc_ubigeo_Receptor

Type struc_ubigeo_Emisor

    idubigeo                        As String
    provincia                       As String
    departamento                    As String
    distrito                        As String
    direccion                       As String

End Type

Global my_struc_ubigeo_Emisor() As struc_ubigeo_Emisor

Type struc_credito

    anticipo                         As String
    fpago                            As String
    serie                            As String
    Numero                           As String
    total                            As String

End Type

Global my_struc_credito() As struc_credito

'fin 20/11/2017 pll

Public Function busca_tipo_comprobante(local1 As String, _
                                       bxtipo As String, _
                                       bxserie As String, _
                                       bxnumero As String, _
                                       acu As String, _
                                       salida As Boolean, _
                                       my_tipo As String, _
                                       my_codcliente As String, _
                                       my_acu As String)

    Dim mysql   As String

    Dim mytable As New ADODB.Recordset

    mysql = ""

    If acu <> "" Then
        mysql = "SELECT tipo,CODIGO " & Chr$(10)
    Else
        mysql = "SELECT tipo,CODIGO,ACU " & Chr$(10)

    End If

    mysql = mysql & "from Factura " & Chr$(10)
    mysql = mysql & "where local='" & "" & local1 & "' " & Chr$(10)
    mysql = mysql & "and serie ='" & "" & bxserie & "' " & Chr$(10)
    mysql = mysql & "and numero='" & "" & bxnumero & "' " & Chr$(10)
 
    If acu <> "" Then
        mysql = mysql & "and acu='" & "" & acu & "' " & Chr$(10)
    Else
        mysql = mysql & "and tipo='" & "" & bxtipo & "' " & Chr$(10)

    End If
 
    mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
 
    If mytable.EOF Then
        salida = False
        Exit Function
    Else
        salida = True

        If acu <> "" Then
            my_tipo = mytable.Fields("TIPO")
            my_codcliente = mytable.Fields("CODIGO")
        Else
            my_tipo = mytable.Fields("TIPO")
            my_codcliente = mytable.Fields("CODIGO")
            my_acu = mytable.Fields("acu")

        End If

    End If

End Function

'este modulo es para crear file txt para enviar sunat Patricia LLamoja 19/04/2017
Function estrae_factura(my_ruc As String, _
                        local1 As String, _
                        bxtipo As String, _
                        bxserie As String, _
                        bxnumero As String, _
                        my_idubigeo As String, _
                        acu As String, _
                        my_struc_datos_empresa() As struc_datos_empresa, _
                        my_struc_ubigeo_Receptor() As struc_ubigeo_Receptor, _
                        my_carga_busca_cliente() As struc_busca_cliente, _
                        my_struc_credito() As struc_credito, _
                        my_struc_ubigeo_Emisor() As struc_ubigeo_Emisor, _
                        my_struc_Etransporte() As struc_Etransporte, _
                        file As String)
         
    Dim mysql           As String

    Dim mytable         As New ADODB.Recordset

    Dim hastaCuanto     As Integer

    Dim nuevoDato       As String

    Dim myDato          As String

    Dim my_precioSinigv As Currency

    Dim myREG           As String

    Dim xnumero         As String

    Dim xserie          As String

    mysql = ""
    'Datos de la Boleta de Venta (Es identico a la factura)
    mysql = "SELECT d.impuesto as impuestod, d.igv as igvd, f.SERIE," & Chr$(10)
    mysql = mysql & "f.NUMERO, " & Chr$(10) 'serie y número correlativo item1
    mysql = mysql & "CONVERT(VARCHAR(10),f.FECHA, 120) AS FECHA ," & Chr$(10) 'Fecha de emisión item 3
    mysql = mysql & "CONVERT(VARCHAR(10),f.FECHAE, 120) AS FECHAE ," & Chr$(10) 'Fechavencimiento item 4
    mysql = mysql & "f.tipo," & Chr$(10) 'TipoDocumento item 3
    mysql = mysql & "Case f.moneda" & Chr$(10) ' TipoMoneda item 4
    mysql = mysql & "WHEN 'S' THEN 'PEN'" & Chr$(10) ' TipoMoneda item 4
    mysql = mysql & "WHEN 'D' THEN 'USD'" & Chr$(10) ' TipoMoneda item 4
    mysql = mysql & "END AS MONEDA," & Chr$(10) ' TipoMoneda item 4
    '****aqui todo con parametrizacion va al txt
    'Documentos de referencia
    mysql = mysql & "f.adetotal," & Chr$(10) 'serie item 15 'adelanto total

    ' Campos Adicionales FE 19/05/2018
    mysql = mysql & "f.observa," & Chr$(10)
    ' Campos Adicionales FE 19/05/2018

    mysql = mysql & "f.serie1," & Chr$(10) 'serie item 16
    mysql = mysql & "f.numero1," & Chr$(10) 'numero item 16
    mysql = mysql & "f.tipo1," & Chr$(10) 'tipo item 16
    mysql = mysql & "f.tdetra," & Chr$(10) 'tipo item 18
    'aqui los datos de partida y llegada
    mysql = mysql & "f.partida," & Chr$(10)
    mysql = mysql & "f.destino," & Chr$(10)
    'Datos del detalle o Ítem de la Factura
    mysql = mysql & "d.unidad," & Chr$(10)  'UnidadMedidaItem
    mysql = mysql & "d.cantidad," & Chr$(10)  'CantidadItem
    mysql = mysql & "d.producto," & Chr$(10) 'Producto
    mysql = mysql & "d.descripcio as descripcioP," & Chr$(10) 'DescripcionItem PRODUCTO
    mysql = mysql & "isnull(p.costoinisigv,0) as costoinisigv," & Chr$(10) 'ValorUnitarioSinIgv
    mysql = mysql & "d.precio," & Chr$(10) 'PrecioUnitarioConIgv
    'CodTipoPrecioVtaUnitarioItem 01 precio unitario, 02 valor referencial txt
    mysql = mysql & "f.subtotal," & Chr$(10) 'ImporteIGVItem
    'CodigoAfectacionIGVItem **Afectación al IGV - Catálogo No. 07 txt
    mysql = mysql & "f.tisc," & Chr$(10) 'MontoISCItem
    mysql = mysql & "d.tisc as dtisc," & Chr$(10) 'MontoISCItemXDetalle
    
    'para la posicon rojo 6
    mysql = mysql & "isnull(p.costou,0) as costou," & Chr$(10) 'ValorVentaItem
    'para la posicion 8
    mysql = mysql & "f.impuesto," & Chr$(10) 'ValorVentaItem
    mysql = mysql & "p.igv," & Chr$(10) 'posicion 8
    'mysql = mysql & "d.descuento, " & Chr$(10) 'DescuentoItem
    mysql = mysql & "f.descuento as tdescuento, " & Chr$(10) 'DescuentoItem
    mysql = mysql & "d.descuento as descuento, " & Chr$(10) 'DescuentoItem
    mysql = mysql & "d.cantidad * d.precio as ValorVenta," & Chr$(10) 'ValorVentaItem" y la deduccion de los descuentos
    'Totales de la Boleta de Venta  (Es identico a la Factura) ***
    mysql = mysql & "f.total," & Chr$(10) 'TotalOPGravadas item 30
    'TotalOPExoneradas item 32 valor venta no incluye isc aqui **se calcula text
    'TotalOPGratuita 'tem 33 se utiliza si es gratuita **se calcula text
    'TotalDescuentos item 34 si hubo descuentos **se calcula text
    mysql = mysql & "d.subtotal," & Chr$(10) 'SumatoriaIGV item 35 **se calcula text
    mysql = mysql & "d.comision as comision," & Chr$(10)
    mysql = mysql & "f.impuesto as impuesto," & Chr$(10)
    mysql = mysql & "f.subtotal as subtotal," & Chr$(10)
    mysql = mysql & "f.neto as neto," & Chr$(10)
    mysql = mysql & "f.gravado as gravado," & Chr$(10)
    '***Información adicional - Percepciones
    mysql = mysql & "f.percepcion as percepcion ," & Chr$(10) 'BaseImponiblePercepcion 41
    mysql = mysql & "d.tpercepcio," & Chr$(10) ' MontoPercepcion 41 *aqui calcular txt
    '***Información Adicional - Factura Guía /Marca y placa de Vehiculo item 52
    '14/02/2018 pll

    ' Testing Proyecto Facturacion Electronica
    'mysql = mysql & "fp.fpago" & Chr$(10)
    mysql = mysql & "f.fpago as fpago," & Chr$(10)
    mysql = mysql & "p.servicio as pservicio" & Chr$(10) ' MontoPercepcion 41 *aqui calcular txt
    ' Testing Proyecto Facturacion Electronica

    mysql = mysql & "from factura f," & Chr$(10)
    mysql = mysql & "detalle d," & Chr$(10)
    mysql = mysql & "producto p" & Chr$(10)
    mysql = mysql & "where D.dua IS NULL AND f.local='" & "" & local1 & "' " & Chr$(10)
    mysql = mysql & "and f.serie ='" & "" & bxserie & "' " & Chr$(10)
    mysql = mysql & "and f.numero='" & "" & bxnumero & "' " & Chr$(10)
    mysql = mysql & "and f.acu='" & "" & acu & "' " & Chr$(10)
    mysql = mysql & "and f.tipo='2'" & Chr$(10)
    mysql = mysql & "and f.SERIE = d.serie" & Chr$(10)
    mysql = mysql & "and f.NUMERO = d.NUMERO" & Chr$(10)
    mysql = mysql & "and p.producto = d.PRODUCTO" & Chr$(10)
    mysql = mysql & "and f.tipo = d.tipo" & Chr$(10)

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
    'D.dua IS NULL
    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    mytable.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytable.RecordCount > 0 Then
        'para el numero
        hastaCuanto = 8 - Len(mytable.Fields("NUMERO"))
        myDato = mytable.Fields("NUMERO")
        Call E_llenar_zero(hastaCuanto, myDato, xnumero)
        'PARA LA SERIE
 
        Call E_llenar_zero(4 - Len(mytable.Fields("SERIE")), mytable.Fields("SERIE"), xserie)
        FileName = "D:\ce_output\CREA\" & my_ruc & "_01" & "_" & xserie & "-" & xnumero & ".INPUT.TXT"
        file = my_ruc & "_01" & "_" & xserie & "-" & xnumero & ".INPUT.TXT"
 
        Do
            c = c + 1 'el contador sirve para la lista

            If mytable.EOF Then Exit Do
            'aqui es para abrir el file txt y llenarlo
            Filelibero1 = FreeFile
            Open FileName For Append As #Filelibero1
            conta = conta + 1
            conta_record = mytable.RecordCount

            'CABECERA
            If c = 1 Then
                myREG = myREG & "H"
                myREG = myREG & "|"
      
                'posicion 1 Serie/numero
                hastaCuanto = 4 - Trim(Len(mytable.Fields("SERIE")))
                Call E_llenar_zero(3 - Len(mytable.Fields("SERIE")), mytable.Fields("SERIE"), nuevoDato)
                myREG = myREG & "" & Trim(nuevoDato)
                myREG = myREG & "-"

                If Trim(mytable.Fields("NUMERO")) = "vacio" Then
                    myREG = myREG & "|"
                Else
                    hastaCuanto = 8 - Trim$(Len(mytable.Fields("NUMERO")))
                    myDato = mytable.Fields("NUMERO")
                    Call E_llenar_zero(hastaCuanto, myDato, nuevoDato)
                    myREG = myREG & nuevoDato
                    myREG = myREG & "|"

                End If

                'posicion 2 moneda
                If Trim(mytable.Fields("MONEDA")) = Null Then
                    myREG = myREG & "|"
                Else
                    Call llenar_datos(3 - Trim(Len(mytable.Fields("MONEDA"))), mytable.Fields("MONEDA"), nuevoDato)
                    myREG = myREG & nuevoDato
                    myREG = myREG & "|"

                End If

                'posicion 3 Fecha de emisión
                If Trim(mytable.Fields("FECHA")) = 0 Then
                    myREG = myREG & "|"
                Else
                    hastaCuanto = 10 - Trim(Len(mytable.Fields("FECHA")))
                    myDato = mytable.Fields("FECHA")
                    Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                    myREG = myREG & nuevoDato
                    myREG = myREG & "|"

                End If
    
                'posicion 4 fecha vencimiento
                If Trim(mytable.Fields("FECHAE")) = 0 Then
                    myREG = myREG & "|"
                Else
                    hastaCuanto = 10 - Trim(Len(mytable.Fields("FECHAE")))
                    myDato = mytable.Fields("FECHAE")
                    Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                    myREG = myREG & nuevoDato
                    myREG = myREG & "|"

                End If
                
                'posicion 5 OrdenCompra item
                myREG = myREG & "|"
      
                ' factura de exportacion
                'posicion 6 TipoDocIdentidadReceptor
                If Len(my_carga_busca_cliente(0).RUC) = 11 Then
                    myREG = myREG & "6" 'REG. UNICO DE CONTRIBUYENTES
                    myREG = myREG & "|"
                ElseIf Len(my_carga_busca_cliente(0).RUC) = 8 Then
                    myREG = myREG & "1" 'DOC. NACIONAL DE IDENTIDAD
                    myREG = myREG & "|"
                ElseIf Len(my_carga_busca_cliente(0).RUC) = 9 Then
                    myREG = myREG & "4" 'CARNET DE EXTRANJERIA
                    myREG = myREG & "|"
                Else
                    myREG = myREG & "0" 'OTROS
                    myREG = myREG & "|"

                End If

                ' factura de exportacion
    
                'posicion 7  NumeroDocIdentidadReceptor
                Call E_llenar_datos(11 - Trim(Len(my_carga_busca_cliente(0).RUC)), Trim(my_carga_busca_cliente(0).RUC), nuevoDato) 'posicion 7
                myREG = myREG & nuevoDato
                myREG = myREG & "|"
         
                'posicion 8  RazonSocialReceptor
                Call E_llenar_datos(100 - Len(my_carga_busca_cliente(0).nombre), my_carga_busca_cliente(0).nombre, nuevoDato) 'RazonSocialReceptor 'posicion 8
                myREG = myREG & Trim(nuevoDato)
                myREG = myREG & "|"
     
                'posicion 9 CorreoReceptor
                If my_carga_busca_cliente(0).correo <> "" Then
                    Call E_llenar_datos(100 - Len(my_carga_busca_cliente(0).correo), my_carga_busca_cliente(0).correo, nuevoDato)
                    myREG = myREG & Trim(nuevoDato)
                    myREG = myREG & "|"
                Else
                    myREG = myREG & "|"

                End If
 
                'posicion 10 TotalOPGravadas" aqui es el subtotal sin igv
                Call E_llenar_datos(15 - Len(mytable.Fields("subtotal")), mytable.Fields("subtotal"), nuevoDato)
                nuevoDato = nuevoDato - mytable.Fields("gravado")
                myREG = myREG & Format(Trim(nuevoDato), "0.00")
                myREG = myREG & "|"

                'posicion 11 TotalOPNoGravadas
                If my_struc_datos_empresa(0).Toperacion = "I" Then
                    Call E_llenar_datos(15 - Len(mytable.Fields("gravado")), mytable.Fields("gravado"), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|"
                Else
                    myREG = myREG & "0.00|"

                End If
     
                'posicion 12  TotalOPExoneradas
                If my_struc_datos_empresa(0).Toperacion = "E" Then
                    Call E_llenar_datos(15 - Len(mytable.Fields("gravado")), mytable.Fields("gravado"), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|"
                Else
                    myREG = myREG & "0.00|"

                End If
      
                'posicion 13 TotalOPGratuitas
                myREG = myREG & "0.00|"
           
                'posicion 14 TotalDescuentos
                If Trim("" & mytable.Fields("tdescuento")) = "0" Then
                    myREG = myREG & "0.00|"
                Else
                    myREG = myREG & Format(Trim("" & mytable.Fields("tdescuento")), "0.00")
                    myREG = myREG & "|"

                End If
 
                'posicion 15 TotalAnticipos
                myREG = myREG & "0.00|"

                ''posicion 16 SumatoriaIGV
                Call E_llenar_datos(15 - Trim(Len(mytable.Fields("impuesto"))), mytable.Fields("impuesto"), nuevoDato)
                myREG = myREG & Format(Trim(nuevoDato), "0.00")
                myREG = myREG & "|"

                'posicion 17 SumatoriaISC
                Call E_llenar_datos(15 - Trim(Len("" & mytable.Fields("tisc"))), Trim("" & mytable.Fields("tisc")), nuevoDato)
                myREG = myREG & Format(Trim(nuevoDato), "0.00")
                myREG = myREG & "|"

                'posicion 18 'SumatoriaOtrosTributos
                myREG = myREG & "0.00|"
    
                'posicion 19 DescuentosGlobales
                If Trim("" & mytable.Fields("tdescuento")) = "0" Then
                    myREG = myREG & "0.00|"
                Else
                    Call E_llenar_datos(15 - Trim(Len("" & mytable.Fields("tdescuento"))), "" & mytable.Fields("tdescuento"), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|"

                End If

                'posicion 20 SumatoriaOtrosCargos
                myREG = myREG & "0.00|"
                
                'posicion 21 ImporteTotalVenta
                Call E_llenar_datos(15 - Trim(Len(mytable.Fields("total"))), mytable.Fields("total"), nuevoDato)
                myREG = myREG & Format(Trim(Round(nuevoDato, 2)), "0.00")
                myREG = myREG & "|"
    
                'posicion 22 MontoEnLetras/no es obligatorio
                myREG = myREG & "|"
      
                'posicion 23 BaseImponiblePercepcion
                Call E_llenar_datos(15 - Trim(Len(mytable.Fields("total"))), mytable.Fields("total"), nuevoDato)
                myREG = myREG & Format(Trim(nuevoDato), "0.00")
                myREG = myREG & "|"
    
                'posicion 24 MontoPercepcion
                If mytable.Fields("tpercepcio") <> "" Then
                    Call E_llenar_datos(15 - Trim(Len(mytable.Fields("tpercepcio"))), mytable.Fields("tpercepcio"), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|"
                Else
                    myREG = myREG & "|"

                End If

                'posicion 25 MontoTotalncluidoPercepcion
                If "" & mytable.Fields("percepcion") = "0" Then
                    myREG = myREG & "0.00|"
                Else
                    Call E_llenar_datos(15 - Trim("" & mytable.Fields("percepcion")), "" & mytable.Fields("percepcion"), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|" 'gion separador

                End If
 
                'posicion 26 PorcentajePercepcion
                myREG = myREG & "0.00|"
                
                'posicion 27 CodigoTipoOperacion
                '''' 17/07/2018 Factura de Exportación
                myREG = myREG & my_tipooperacion & "|"
                '''' 17/07/2018 Factura de Exportación
 
                'posicion 28 CodigoLeyenda catalogo 15 sunat
                myREG = myREG & "|"

                'posicion 29 DescripcionLeyenda catalogo 15 sunat
                myREG = myREG & "|"
    
                'posicion 30  UbigeoReceptor/catalogo 49
                If my_struc_ubigeo_Receptor(0).idubigeo = "" Then
                    myREG = myREG & "0" & "|"
                Else
                    Call E_llenar_datos(100 - Len(my_struc_ubigeo_Receptor(0).idubigeo), my_struc_ubigeo_Receptor(0).idubigeo, nuevoDato)
                    myREG = myREG & Trim(nuevoDato)
                    myREG = myREG & "|"

                End If

                'posicion 31 DireccionReceptor
                If my_carga_busca_cliente(0).direccion <> "" Then
                    Call E_llenar_datos(100 - Len(my_struc_ubigeo_Receptor(0).direccion), my_carga_busca_cliente(0).direccion, nuevoDato)
                    myREG = myREG & Trim(nuevoDato)
                    myREG = myREG & "|"
                Else
                    myREG = myREG & "|"

                End If

                'posicion 32 UrbanizacionPuntoLlegada
                myREG = myREG & "|"

                'posicion 33 ProvinciaReceptor
                If my_carga_busca_cliente(0).provincia <> "" Then
                    Call E_llenar_datos(5 - Len(my_carga_busca_cliente(0).provincia), my_carga_busca_cliente(0).provincia, nuevoDato) 'item 50 /posicion 52
                    myREG = myREG & Trim(nuevoDato)
                    myREG = myREG & "|"
                Else
                    myREG = myREG & "|"

                End If

                'posicion 34 DepartamentoReceptor
                If my_struc_ubigeo_Receptor(0).departamento <> "" Then
                    Call E_llenar_datos(5 - Len(my_struc_ubigeo_Receptor(0).departamento), my_struc_ubigeo_Receptor(0).departamento, nuevoDato) 'item 50 /posicion 52
                    myREG = myREG & Trim(nuevoDato)
                    myREG = myREG & "|"
                Else
                    myREG = myREG & "|"

                End If

                'posicion 35 DistritoReceptor
                If my_struc_ubigeo_Receptor(0).distrito <> "" Then
                    Call E_llenar_datos(5 - Len(my_struc_ubigeo_Receptor(0).distrito), my_struc_ubigeo_Receptor(0).distrito, nuevoDato) 'item 50 /posicion 52
                    myREG = myREG & Trim(nuevoDato)
                    myREG = myREG & "|"
                Else
                    myREG = myREG & "|"

                End If
 
                'posicion 36 CodigoPaisReceptor
                myREG = myREG & "PE|"
                  
                ' posicion 37 SerieNumeroGuia
                If my_struc_Etransporte(0).partida <> "" Then
                    hastaCuanto = 4 - Trim(Len(mytable.Fields("SERIE1")))
                    Call E_llenar_zero(3 - Len(mytable.Fields("SERIE1")), mytable.Fields("SERIE1"), nuevoDato)
                    myREG = myREG & "" & Trim(nuevoDato)
                    myREG = myREG & "-"

                    If Trim(mytable.Fields("NUMERO1")) = "vacio" Then
                        myREG = myREG & "|"
                    Else
                        hastaCuanto = 8 - Trim$(Len(mytable.Fields("NUMERO1")))
                        myDato = mytable.Fields("NUMERO1")
                        Call E_llenar_zero(hastaCuanto, myDato, nuevoDato)
                        myREG = myREG & nuevoDato
                        myREG = myREG & "|"

                    End If
  
                    'posicion 38 TipoDocumentoGuia
                    myREG = myREG & "01|"

                    'posicion 39 NumeroDocumentoRelacionado p
                    If mytable.Fields("numero") <> "" Then
                        Call E_llenar_datos(4 - Trim(Len(mytable.Fields("numero"))), mytable.Fields("numero"), nuevoDato)
                        myREG = myREG & "H" & Trim(nuevoDato)
                        myREG = myREG & "-"
                    Else
                        myREG = myREG & "|"

                    End If

                    ' posicion 40 TipoDocumentoRelacionado
                    If mytable.Fields("tipo") <> "" Then
                        myREG = myREG & "03|" 'factura

                    End If

                Else
                    myREG = myREG & "||||"

                End If

                'posicion 41 TotalRetencion
                myREG = myREG & "0.00|"
                'posicion 42 PorcentajeRetencion
                myREG = myREG & "0.00|"
                'posicion 43 DescripcionDetraccion
                myREG = myREG & "|"
                'posicion 44 TotalBonificacion
                myREG = myREG & "0.00|"
                
                'posicion 45 UbigeoLugarEntrega
                If my_struc_ubigeo_Receptor(0).idubigeo <> "" Then
                    Call E_llenar_datos(5 - Len(my_struc_ubigeo_Receptor(0).idubigeo), my_struc_ubigeo_Receptor(0).idubigeo, nuevoDato) 'item 50 /posicion 52
                    myREG = myREG & Trim(nuevoDato)
                    myREG = myREG & "|"
                Else
                    myREG = myREG & "|"

                End If
    
                'posicion 46 DireccionLugarEntrega
                If mytable.Fields("partida") <> "" Then
                    Call E_llenar_datos(100 - Len(mytable.Fields("partida")), mytable.Fields("partida"), nuevoDato) 'item 51 /posicion 60
                    myREG = myREG & Trim(nuevoDato)
                    myREG = myREG & "|"
                Else
                    myREG = myREG & "|"

                End If
                
                'posicion 47 UrbanizacionLugarEntrega
                myREG = myREG & "|"
    
                'posicion 48 ProvinciaLugarEntrega
                myREG = myREG & "|"
 
                'posicion 49 DepartamentoReceptor
                myREG = myREG & "|"
                
                'posicion 50 DistritoReceptor
                myREG = myREG & "|"
  
                'posicion 51 codigoPaisEmisor
                myREG = myREG & "PE|"
 
                'posicion 52 UbigeoPuntoPartida
                If my_struc_ubigeo_Emisor(0).idubigeo <> "" Then
                    Call E_llenar_datos(5 - Len(my_struc_ubigeo_Emisor(0).idubigeo), my_struc_ubigeo_Emisor(0).idubigeo, nuevoDato) 'item 50 /posicion 52
                    myREG = myREG & Trim(nuevoDato)
                    myREG = myREG & "|"
                Else
                    myREG = myREG & "|"

                End If
                
                ' posicion 53 DireccionPuntoPartida
                If mytable.Fields("partida") <> "" Then
                    Call E_llenar_datos(5 - Len(mytable.Fields("partida")), mytable.Fields("partida"), nuevoDato)
                    myREG = myREG & Trim(nuevoDato)
                    myREG = myREG & "|"
                Else
                    myREG = myREG & "|"

                End If

                'posicion 54 UrbanizacionPuntoPartida
                myREG = myREG & "|"
 
                'posicion 55 ProvinciaPuntoPartida
                If my_struc_ubigeo_Emisor(0).provincia <> "" Then
                    Call E_llenar_datos(5 - Len(my_struc_ubigeo_Emisor(0).provincia), my_struc_ubigeo_Emisor(0).provincia, nuevoDato) 'item 50 /posicion 52
                    myREG = myREG & Trim(nuevoDato)
                    myREG = myREG & "|"
                Else
                    myREG = myREG & "|"

                End If
                
                'posicion 56 DepartamentoPuntoPartida
                Call E_llenar_datos(5 - Len(my_struc_ubigeo_Emisor(0).departamento), my_struc_ubigeo_Emisor(0).departamento, nuevoDato) 'item 50 /posicion 52
                myREG = myREG & Trim(nuevoDato)
                myREG = myREG & "|"
                
                'posicion 57 DistritoPuntoPartida
                Call E_llenar_datos(5 - Len(my_struc_ubigeo_Emisor(0).distrito), my_struc_ubigeo_Emisor(0).distrito, nuevoDato) 'item 50 /posicion 52
                myREG = myREG & Trim(nuevoDato)
                myREG = myREG & "|"
                
                'posicion 58CodigoPaisPuntoPartida
                myREG = myREG & "PE"
                myREG = myREG & "|"
                
                'posicion 59 UbigeoPuntoLlegada
                Call E_llenar_datos(5 - Len(my_struc_ubigeo_Receptor(0).idubigeo), my_struc_ubigeo_Receptor(0).idubigeo, nuevoDato) 'item 50 /posicion 52
                myREG = myREG & Trim(nuevoDato)
                myREG = myREG & "|"
 
                'posicion 60 DireccionPuntoLlegada
                Call E_llenar_datos(5 - Len(mytable.Fields("destino")), mytable.Fields("destino"), nuevoDato) 'item 51 /posicion 59
                myREG = myREG & Trim(nuevoDato)
                myREG = myREG & "|"

                'posicion 61 UrbanizacionPuntoLlegada
                myREG = myREG & "|"
 
                'posicion 62 ProvinciaPuntoLlegada
                Call E_llenar_datos(5 - Len(my_struc_ubigeo_Receptor(0).provincia), my_struc_ubigeo_Receptor(0).provincia, nuevoDato) 'item 50 /posicion 52
                myREG = myREG & Trim(nuevoDato)
                myREG = myREG & "|"
 
                ' posicion 63 DepartamentoPuntoLlegada
                Call E_llenar_datos(5 - Len(my_struc_ubigeo_Receptor(0).departamento), my_struc_ubigeo_Receptor(0).departamento, nuevoDato) 'item 50 /posicion 52
                myREG = myREG & Trim(nuevoDato)
                myREG = myREG & "|"
 
                'posicion 64 DistritoPuntoLlegada
                Call E_llenar_datos(5 - Len(my_struc_ubigeo_Receptor(0).distrito), my_struc_ubigeo_Receptor(0).distrito, nuevoDato) 'item 50 /posicion 52
                myREG = myREG & Trim(nuevoDato)
                myREG = myREG & "|"
                 
                'posicion 65 CodigoPaisPuntoLlegada
                myREG = myREG & "PE"
                myREG = myREG & "|"

                'posicion 66 PlacaVehiculo
                If my_struc_Etransporte(0).placa <> "" Then
                    Call E_llenar_datos(8 - Trim(Len(my_struc_Etransporte(0).placa)), my_struc_Etransporte(0).placa, nuevoDato)
                    myREG = myREG & Trim(nuevoDato)
                    myREG = myREG & "|"
                Else
                    myREG = myREG & "|"

                End If

                'posicion 67 NumeroCertificadoVehicular
                If my_struc_Etransporte(0).licencia <> "" Then
                    Call E_llenar_datos(30 - Len(my_struc_Etransporte(0).licencia), my_struc_Etransporte(0).licencia, nuevoDato)
                    myREG = myREG & Trim(nuevoDato)
                    myREG = myREG & "|"
                Else
                    myREG = myREG & "|"

                End If

                'posicion 68 MarcaVehiculo
                If my_struc_Etransporte(0).marca <> "" Then
                    Call E_llenar_datos(30 - Trim(Len(my_struc_Etransporte(0).marca)), my_struc_Etransporte(0).marca, nuevoDato)
                    myREG = myREG & Trim(nuevoDato)
                    myREG = myREG & "|"
                Else
                    myREG = myREG & "|"

                End If
                
                'posicion 69 NumeroLicenciaConducir
                If my_struc_Etransporte(0).licencia <> "" Then
                    Call E_llenar_datos(30 - Trim(Len(my_struc_Etransporte(0).licencia)), my_struc_Etransporte(0).licencia, nuevoDato)
                    myREG = myREG & Trim(nuevoDato)
                    myREG = myREG & "|"
                Else
                    myREG = myREG & "|"

                End If
                
                'posicion 70 RucTransportista
                If my_struc_Etransporte(0).RUC <> "" Then
                    myREG = myREG & "01" & "|"
                Else
                    myREG = myREG & "|"

                End If

                'posicion 71 CodigoDocdentidadTransportista
                myREG = myREG & "|"

                'posicion 72 RazonSocialTransportista
                If my_struc_Etransporte(0).nombreT <> "" Then
                    Call E_llenar_datos(100 - Len(mytable.Fields("nombre")), mytable.Fields("nombre"), nuevoDato)
                    myREG = myREG & Trim(nuevoDato)
                    myREG = myREG & "|"
                Else
                    myREG = myREG & "|"

                End If

                'posicion 73 ModalidadTransporte
                myREG = myREG & "01|" 'privado tabla sunat
                myREG = myREG & "|"

                'posicion 74 TotlaPesoBruto
                myREG = myREG & "|"
    
                'posicion 75 UnidadMedidaPesoBruto
                myREG = myREG & "|"
   
                'posicion 76 PlacaVehiculoGastoRenta
                myREG = myREG & "|"
    
                'posicion 77 DTCodigoBienServicio
                myREG = myREG & "|"
    
                'posicion 77 DTNumeroCuentaBancoNacion
                myREG = myREG & "|"
     
                'posicion 77 DTPorcentajeDetraccion
                myREG = myREG & "|"
    
                'posicion 77 **Detracciones/monto de la detraccion
                myREG = myREG & "|"

                'posicion 77 **Detracciones - Recursos Hidrobiológicos
                '**Detracciones/Nombre y matrícula de la embarcación pesquera utilizada
                myREG = myREG & "|"
    
                'posicion 77 **Detracciones/Descripción del tipo y cantidad de la especie vendida
                myREG = myREG & "|"
    
                'posicion 77 **Detracciones/Lugar de la descarga
                myREG = myREG & "|"
    
                'posicion 77 **Detracciones/Fecha de la descarga
                myREG = myREG & "|"
    
                'posicion 77 ***En caso de mas de un recurso hidrobiológico/Matricula de embarcación
                myREG = myREG & "|"
   
                'posicion 77 ***En caso de mas de un recurso hidrobiológico/Nombre de la embarcación
                myREG = myREG & "|"
    
                'posicion 77 **En caso de mas de un recurso hidrobiológico/Descripción del tipo y cantidad de la especie vendida
                myREG = myREG & "|"
    
                'posicion 77 **En caso de mas de un recurso hidrobiológico/Lugar de la descarga
                myREG = myREG & "|"
    
                'posicion 77 **En caso de mas de un recurso hidrobiológico/Fecha de descarga
                myREG = myREG & "|"
    
                'posicion 77 **Detracciones - Servicio de Transporte/Valor referencial del servicio
                'DTCodigoConceptosTributarios
                myREG = myREG & "|"
    
                'posicion 77 **Detracciones - Servicio de Transporte/Numero de registro MTC
                'DTValorRefServicioTransporte
                myREG = myREG & "|"
                
                'posicion 77 **Detracciones - Servicio de Transporte/Numero de registro MTC
                'DTNumeroRegistroMTC
                myREG = myREG & "|"
    
                'posicion 77 **Detracciones - Servicio de Transporte/Configuracion vehicular
                'DTConfiguracionVehicular
                myREG = myREG & "|"
    
                'posicion 77 **Detracciones - Servicio de Transporte/Punto de destino
                'DTPuntoOrigen
                myREG = myREG & "|"

                'posicion 77 **Detracciones - Servicio de Transporte/Punto de destino
                'DTPuntoOrigen
                myREG = myREG & "|"

                'posicion 77 **Detracciones - Servicio de Transporte/DTDescripcionViaje
                'DTDescripcionViaje
                myREG = myREG & "|"
    
                'posicion 77 **Detracciones - Servicio de Transporte/DTValorReferencialPreliminar
                'DTValorReferencialPreliminar
                myREG = myREG & "|"
    
                'posicion 77 **En caso de detallar el concepto del viaje/ Monto Refrencial
                'DTMontoRefViaje
                myREG = myREG & "|"
 
                'posicion 77 **En caso de detallar el concepto del viaje/Monto Refrencial
                'DTMonedaMontoRefViaje
                myREG = myREG & "|"
    
                'posicion 77 **En caso de detallar el concepto del viaje/Monto Referencial Preliminar por Viaje
                'DTMontoRefPreliminarViaje
                myREG = myREG & "|"
    
                'posicion 77 **En caso de detallar el concepto del viaje/Indica factor de retorno de viaje
                'DTIndFactorRetViaje
                myREG = myREG & "|"
    
                'posicion 77 **En caso de detallar el concepto del viaje/Punto de origen del viaje
                'DTPuntoOrigenViaje
                myREG = myREG & "|"
    
                'posicion 77 **En caso de detallar el concepto del viaje/Punto de destino del viaje
                'DTPuntoDestinoViaje
                myREG = myREG & "|"
    
                'posicion 77 **En caso de detallar el concepto del viaje/Carga efectiva en Tn por vehículo
                'DTCargaEfectivaVehiculoViaje
                myREG = myREG & "|"
    
                'posicion 77 **En caso de detallar el concepto del viaje/Carga efectiva en Tn por vehículo
                'DTUnidadMedCargaEfectivaViaje
                myREG = myREG & "|"
    
                'posicion 77 **En caso de detallar el concepto del viaje/Monto Referencial por Vehículo
                'DTMontoRefVehiculoViaje
                myREG = myREG & "|"
    
                'posicion 77 **En caso de detallar el concepto del viaje/Monto Referencial por Vehículo
                'DTMonedaMontoRefVehiculoViaje
                myREG = myREG & "|"
    
                'posicion 77 **En caso de detallar el concepto del viaje/Configuracion vehicular del vehículo
                'DTConfiguracionVehicularViaje
                myREG = myREG & "|"
    
                'posicion 77 **En caso de detallar el concepto del viaje/Carga Util en Tn del vehículo en viaje
                'DTCargaUtilVehiculoViaje
                myREG = myREG & "|"
    
                'posicion 77 ***En caso de detallar el concepto del viajes/Carga Util en Tn del vehículo en viaje
                'DTUnidadMedCargaUtilViaje
                myREG = myREG & "|"
    
                'posicion 77 ***En caso de detallar el concepto del viaje/Monto Referencial por TM en viaje
                'DTMontoRefTMViaje
                myREG = myREG & "|"
    
                'posicion 77 **Beneficio - Establecimiento de Hospedajes/Código País de emisión del pasaporte
                'FechaEmisionPasaporte
                myREG = myREG & "|"
    
                'posicion 77 **Beneficio - Establecimiento de Hospedajes/Fecha de salida del establecimiento
                'FechaPaisResidenciaNoDomicilio
                myREG = myREG & "|"
    
                'posicion 77 **Beneficio - Establecimiento de Hospedajes/Fecha de ingreso al país
                'FechaIngresoPais
                myREG = myREG & "|"
    
                'posicion 77 **Beneficio - Establecimiento de Hospedajes/Fecha de ingreso al establecimiento
                'FechaIngresoEstablecimiento
                myREG = myREG & "|"
    
                '**Beneficio - Establecimiento de Hospedajes/Fecha de salida del establecimiento
                'FechaSalidaEstablecimiento
                myREG = myREG & "|" 'gion separador muestra de b&H item 92 posicion 116
    
                '**Beneficio - Establecimiento de Hospedajes/Número de días de permanencia
                'FechaDiasPermanencia
                myREG = myREG & "|" 'gion separador muestra de b&H item 93 posicion 117
    
                '**Beneficio - Establecimiento de Hospedajes/Fecha de consumo
                'FechaConsumo
                myREG = myREG & "|" ' muestra de b&H item 94 posicion 118
    
                '**Beneficio - Establecimiento de Hospedajes /Paquete turístico - Nombres y Apellidos del Huésped
                'PaqueteTuristicoNombres
                myREG = myREG & "|" 'gion separador muestra de b&H item 95 posicion 119
    
                '**Beneficio - Establecimiento de Hospedajes /Paquete turístico – Tipo documento identidad del huésped
                ' PaqueteTuristicoDocumento
                myREG = myREG & "|" 'gion separador muestra de b&H item 96 posicion 120
    
                '**Beneficio - Establecimiento de Hospedajes /Numero de documento identidad de huésped
                'NumeroDocumentoHuesped
                myREG = myREG & "|" 'gion separador muestra de b&H item 97 posicion 121
    
                '** Ventas al Sector Público /Numero de Expediente
                'SPNumeroExpediente
                myREG = myREG & "|" 'gion separador muestra de b&H item 98 posicion 122
 
                '** Ventas al Sector Público /Código de unidad ejecutora
                'SPCodigoUnidadEjecutora
                myREG = myREG & "|" 'gion separador  muestra de b&H item 99 posicion 123
    
                '** Ventas al Sector Público /N° de contrato
                'SPNumeroContrato
                myREG = myREG & "|" 'gion separador muestra de b&H item 100 posicion 124
     
                '** Ventas al Sector Público /N° de proceso de selección
                'SPNumeroProcesoSeleccion
                myREG = myREG & "|" 'gion separador muestra de b&H item 101 posicion 125
     
                '**Otra Informacion adicional /FISE (Ley 29852) Fondo Inclusión Social Energético
                'FondoInclusionSocialEnergetico
  
                '''20/02/2018 Kenyo Facturación Electrónica
                'myREG = myREG & "|" 'gion separador muestra de b&H item 105 posicion 126
                'myREG = myREG & "|" 'gion separador muestra de b&H item 105 posicion 126
                '''20/02/2018 Kenyo Facturación Electrónica

                'para la descripcion del tipo de pago 14/02/20180 pll
                'inicio 14/02/2018 pll forma de pago
                If "" & mytable.Fields("fpago") <> "Null" Then
                    If mytable.Fields("fpago") = "3" Then
                        Call E_llenar_datos(30 - Len("CREDITO"), "CREDITO", nuevoDato)
                        myREG = myREG & Trim(nuevoDato)
                    Else
                        Call E_llenar_datos(30 - Len("CONTADO"), "CONTADO", nuevoDato)
                        myREG = myREG & Trim(nuevoDato)

                    End If

                    'myREG = myREG & "|" 'gion separador
                End If

                'fin 14/02/2018 pll forma de pago

                ' Varios Locales FE 18/05/2018
                myDato = my_struc_datos_empresa(0).CodSede
                myREG = myREG & "|"

                If Len(myDato) > 0 Then
                    myREG = myREG & myDato

                End If

                ' Varios Locales FE 18/05/2018
                Print #Filelibero1, myREG
                'para el detalle
 
                myREG = myREG & Chr(13)

            End If

            myREG = ""
            myREG = myREG & "I"
            myREG = myREG & "|"
    
            '**Datos del detalle
            
            'posicion 1 OrdenItem
            myDato = c

            If Len(myDato) = 1 Then
                myREG = myREG & "0" & myDato
                myREG = myREG & "|"
            Else
                myREG = myREG & myDato
                myREG = myREG & "|"

            End If

            'posicion 2 CodigoProductoItem
            If mytable.Fields("producto") = Null Then
                myREG = myREG & "|"
            Else
                Call llenar_datos(30 - Trim(Len(mytable.Fields("producto"))), mytable.Fields("producto"), nuevoDato)
                myREG = myREG & Trim(nuevoDato)
                myREG = myREG & "|"

            End If

            'posicion 3 DescripcionItem
            If mytable.Fields("descripcioP") = Null Then
                myREG = myREG & "|"
            Else
                Call E_llenar_datos(250 - Trim(Len(mytable.Fields("descripcioP"))), mytable.Fields("descripcioP"), nuevoDato)
                myREG = myREG & Trim(nuevoDato)
                myREG = myREG & "|"

            End If
            
            'posicion 4 UnidadMedidaItem
            If mytable.Fields("unidad") = Null Then
                myREG = myREG & "|"
            Else

                If mytable.Fields("pservicio") = "S" Then
                    myREG = myREG & "ZZ|"
                Else
                    myREG = myREG & "NIU|"

                End If

            End If
 
            'posicion 5 CantidadItem
            If mytable.Fields("cantidad") = Null Then
                myREG = myREG & "|"
            Else
                Call E_llenar_datos(23 - Trim(Len(mytable.Fields("cantidad"))), mytable.Fields("cantidad"), nuevoDato)
                myREG = myREG & Trim(nuevoDato)
                myREG = myREG & "|"

            End If

            'posicion 6 ValorVentaItem
            my_precioSinigv = mytable.Fields("precio") / (1 + mytable.Fields("igv") / 100)
            'my_dato = Round(my_precioSinigv, 2) * mytable.Fields("cantidad")
            my_dato = Round(my_precioSinigv, 2) * mytable.Fields("cantidad") - mytable.Fields("descuento")
            myREG = myREG & Format(Trim(my_dato), "0.00")
            myREG = myREG & "|"
    
            'posicion 7 ValorUnitarioSinIgv
            my_precioSinigv = 0
            my_precioSinigv = (mytable.Fields("precio") / (1 + mytable.Fields("igv") / 100))
            my_dato = Round(my_precioSinigv, 2)
            myREG = myREG & Format(Trim(my_dato), "0.00")
            myREG = myREG & "|"
  
            'posicion 8 PrecioUnitarioConIgv
            Call E_llenar_datos(15 - Len(mytable.Fields("precio")), mytable.Fields("precio"), nuevoDato)  'item 28
            myREG = myREG & Format(Trim(nuevoDato), "0.00")
            myREG = myREG & "|" 'gion separador 'es junto con bajo
  
            'posicion 9 DescuentoItem
            If Trim("" & mytable.Fields("descuento")) = "0" Then
                myREG = myREG & "0.00|"
            Else
                Call E_llenar_datos(15 - Trim(Len(mytable.Fields("descuento"))), Trim(mytable.Fields("descuento")), nuevoDato) 'item 28
                myREG = myREG & Format(Trim(nuevoDato), "0.00")
                myREG = myREG & "|"

            End If

            'posicion 10 CodTipoPrecioVtaUnitarioItem
            myREG = myREG & "01"
            myREG = myREG & "|"
            
            'posicion 11 ImporteIGVItem
            '''' 17/07/2018 Factura de Exportación
            Call E_llenar_datos(15 - Trim(Len(mytable.Fields("impuestod"))), mytable.Fields("impuestod"), nuevoDato) 'item 27
            myREG = myREG & Format(Trim(nuevoDato), "0.00")
            myREG = myREG & "|"
            '''' 17/07/2018 Factura de Exportación
        
            'posicion 12 CodigoAfectacionIGVItem
            '''' 17/07/2018 Factura de Exportación
            If mytable.Fields("igvd") = "0" Then
                If my_tipoigv = "40" Then
                    myREG = myREG & "40|"
                Else
                    myREG = myREG & "20|"

                End If

            Else
                myREG = myREG & "10|"

            End If

            '''' 17/07/2018 Factura de Exportación
 
            'posicion 13 MontoISCItem
            If Trim("" & mytable.Fields("dtisc")) = "" Or Trim("" & mytable.Fields("dtisc")) = Null Or Trim("" & mytable.Fields("dtisc")) = "0" Then
                myREG = myREG & "|"
            Else
                Call E_llenar_datos(15 - Trim(Len(mytable.Fields("dtisc"))), Trim(mytable.Fields("dtisc")), nuevoDato)
                myREG = myREG & Format(Trim(nuevoDato), "0.00")
                myREG = myREG & "|"

            End If
            
            'posicion 14 TipoSistemaISCItem
            myREG = myREG & "03|"
       
            'posicion 15 MontoReferencialUnitarioItem
            myREG = myREG & "0.00"
            myREG = myREG & "|"
    
            'posicion 16 CodigoTipoPrecioReferencial
            myREG = myREG & "|"
    
            'posicion 17 MontoReferenciaItem
            myREG = myREG & "|"
          
            '** 'Datos del Emisor 'Ruc item 5
            
            Print #Filelibero1, myREG
   
            ' Campos Adicionales FE 19/05/2018
            If c = conta_record Then
                myREG = myREG & Chr(13)
                myREG = ""
                myREG = myREG & "A"
                myREG = myREG & "|1||OBSERVACIÓN: "
                myREG = myREG & mytable.Fields("observa")
                Print #Filelibero1, myREG

            End If

            '  myREG = myREG & Chr(13)
            '   myREG = ""
            '   myREG = myREG & "A"
            '   myREG = myREG & "|1|100|Orden:ver-4"
            '   Print #Filelibero1, myREG
            '
            '    myREG = myREG & Chr(13)
            '   myREG = ""
            '   myREG = myREG & "A"
            '   myREG = myREG & "|1|100|xxxx"
            '   Print #Filelibero1, myREG
            ' Campos Adicionales FE 19/05/2018
   
            Close #Filelibero1
            mytable.MoveNext
        Loop

    End If

    mytable.Close
    'aqui copia el archivo txt

    origen = FileName

    destino = "D:\ce_Input\" & file

    ' Testing Proyecto Facturacion Electronica 28/02/2018
    'FileCopy origen, destino
    Dim fso As New Scripting.FileSystemObject

    fso.MoveFile origen, destino
    ' Testing Proyecto Facturacion Electronica  28/02/2018

End Function

Function estrae_boleta(my_ruc As String, _
                       local1 As String, _
                       bxtipo As String, _
                       bxserie As String, _
                       bxnumero As String, _
                       my_idubigeo As String, _
                       acu As String, _
                       my_struc_datos_empresa() As struc_datos_empresa, _
                       my_struc_ubigeo_Receptor() As struc_ubigeo_Receptor, _
                       my_carga_busca_cliente() As struc_busca_cliente, _
                       my_struc_credito() As struc_credito, _
                       my_struc_ubigeo_Emisor() As struc_ubigeo_Emisor, _
                       file As String)
                     
    Dim mysql           As String

    Dim mytable         As New ADODB.Recordset

    Dim hastaCuanto     As Integer

    Dim nuevoDato       As String

    Dim myDato          As String

    Dim my_precioSinigv As Currency

    Dim myREG           As String

    Dim xnumero         As String

    Dim xserie          As String
   
    mysql = ""
    'Datos de la Boleta de Venta (Es identico a la factura)
    mysql = "SELECT d.impuesto as impuestod, d.igv as igvd, f.SERIE," & Chr$(10)
    mysql = mysql & "f.NUMERO, " & Chr$(10) 'serie y número correlativo item1
    mysql = mysql & "CONVERT(VARCHAR(10),f.FECHA, 120) AS FECHA ," & Chr$(10) 'Fecha de emisión item 3
    mysql = mysql & "f.tipo," & Chr$(10) 'TipoDocumento item 3
    mysql = mysql & "Case f.moneda" & Chr$(10) ' TipoMoneda item 4
    mysql = mysql & "WHEN 'S' THEN 'PEN'" & Chr$(10) ' TipoMoneda item 4
    mysql = mysql & "WHEN 'D' THEN 'USD'" & Chr$(10) ' TipoMoneda item 4
    mysql = mysql & "END AS MONEDA," & Chr$(10) ' TipoMoneda item 4
    '****aqui todo con parametrizacion va al txt
    'Documentos de referencia

    ' Campos Adicionales FE 19/05/2018
    mysql = mysql & "f.observa," & Chr$(10)
    ' Campos Adicionales FE 19/05/2018
  
    mysql = mysql & "f.adetotal," & Chr$(10) 'serie item 15 'adelanto total
    mysql = mysql & "f.serie1," & Chr$(10) 'serie item 16
    mysql = mysql & "f.numero1," & Chr$(10) 'numero item 16
    mysql = mysql & "f.tipo1," & Chr$(10) 'tipo item 16
    mysql = mysql & "f.tdetra," & Chr$(10) 'tipo item 18
    'aqui los datos de partida y llegada
    mysql = mysql & "f.partida," & Chr$(10)
    mysql = mysql & "f.destino," & Chr$(10)
    'Datos del detalle o Ítem de la Factura
    mysql = mysql & "d.unidad," & Chr$(10)  'UnidadMedidaItem
    mysql = mysql & "d.cantidad," & Chr$(10)  'CantidadItem
    mysql = mysql & "d.producto," & Chr$(10) 'Producto
    mysql = mysql & "d.descripcio as descripcioP," & Chr$(10) 'DescripcionItem producto
    mysql = mysql & "d.total as dtotal," & Chr$(10) 'DescripcionItem
    mysql = mysql & "isnull(p.costoinisigv,0) as costoinisigv," & Chr$(10) 'ValorUnitarioSinIgv
    mysql = mysql & "d.precio," & Chr$(10) 'PrecioUnitarioConIgv
    'CodTipoPrecioVtaUnitarioItem 01 precio unitario, 02 valor referencial txt
    mysql = mysql & "f.subtotal," & Chr$(10) 'ImporteIGVItem
    'CodigoAfectacionIGVItem **Afectación al IGV - Catálogo No. 07 txt
    mysql = mysql & "f.tisc as tisc," & Chr$(10) 'MontoISCItem
    mysql = mysql & "d.tisc as dtisc," & Chr$(10) 'MontoISCItemXDetalle
    
    'para la posicon rojo 6
    mysql = mysql & "isnull(p.costou,0) as costou," & Chr$(10) 'ValorVentaItem
    'para la posicion 8
    mysql = mysql & "f.impuesto," & Chr$(10) 'ValorVentaItem
    mysql = mysql & "p.igv," & Chr$(10) 'posicion 8
    mysql = mysql & "f.descuento as tdescuento, " & Chr$(10) 'DescuentoItem
    mysql = mysql & "d.descuento as descuento, " & Chr$(10) 'DescuentoItem
    mysql = mysql & "d.cantidad * d.precio as ValorVenta," & Chr$(10) 'ValorVentaItem" y la deduccion de los descuentos
    'Totales de la Boleta de Venta  (Es identico a la Factura) ***
    mysql = mysql & "f.total," & Chr$(10) 'TotalOPGravadas item 30
    mysql = mysql & "isnull(d.subtotal,0) as subtotald, " & Chr$(10) 'SumatoriaIGV item 35 **se calcula text
    mysql = mysql & "d.Comision as comision," & Chr$(10)
    '**Posicion 10/CodTipoPrecioVtaUnitarioItem 01 incluye igv 02 operaciones no honerosas
    mysql = mysql & "f.impuesto as impuesto," & Chr$(10)
    mysql = mysql & "f.subtotal as subtotal," & Chr$(10)
    mysql = mysql & "f.neto as neto," & Chr$(10)
    mysql = mysql & "f.gravado as gravado," & Chr$(10)
    '***Información adicional - Percepciones
    mysql = mysql & "f.percepcion," & Chr$(10) 'BaseImponiblePercepcion 41
    mysql = mysql & "d.tpercepcio," & Chr$(10) ' MontoPercepcion 41 *aqui calcular txt
    '***Información Adicional - Factura Guía /Marca y placa de Vehiculo item 52
    '14/02/2018 pll

    ' Testing Proyecto Facturacion Electronica
    'mysql = mysql & "fp.fpago" & Chr$(10)
    mysql = mysql & "p.servicio as pservicio, " & Chr$(10) ' MontoPercepcion 41 *aqui calcular txt
    ' Testing Proyecto Facturacion Electronica
    mysql = mysql & "f.fpago  as fpago" & Chr$(10) ' condicion de pago

    mysql = mysql & "from factura f," & Chr$(10)
    mysql = mysql & "detalle d," & Chr$(10)
    mysql = mysql & "producto p" & Chr$(10)
    mysql = mysql & "where D.DUA IS NULL AND f.local='" & "" & local1 & "' " & Chr$(10)
    mysql = mysql & "and f.serie ='" & "" & bxserie & "' " & Chr$(10)
    mysql = mysql & "and f.numero='" & "" & bxnumero & "' " & Chr$(10)
    mysql = mysql & "and f.acu='" & "" & acu & "' " & Chr$(10)
    mysql = mysql & "and f.tipo='1'" & Chr$(10)
    mysql = mysql & "and f.SERIE = d.serie" & Chr$(10)
    mysql = mysql & "and f.NUMERO = d.NUMERO" & Chr$(10)
    mysql = mysql & "and p.producto = d.PRODUCTO" & Chr$(10)
    mysql = mysql & "and f.tipo = d.tipo" & Chr$(10)

    ' Testing Proyecto Facturacion Electronica 18/05/2018
    'Se arregla forma de pago. Se quita de la query
    ' Testing Proyecto Facturacion Electronica 18/05/2018

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
    'D.dua IS NULL
    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    mytable.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytable.RecordCount > 0 Then

        hastaCuanto = 8 - Len(mytable.Fields("NUMERO"))
        myDato = mytable.Fields("NUMERO")
        Call E_llenar_zero(hastaCuanto, myDato, xnumero)

        Call E_llenar_zero(4 - Len(mytable.Fields("SERIE")), mytable.Fields("SERIE"), xserie)
        FileName = "D:\ce_output\CREA\" & my_ruc & "_03" & "_" & xserie & "-" & xnumero & ".INPUT.TXT"
        file = my_ruc & "_03" & "_" & xserie & "-" & xnumero & ".INPUT.TXT"
  
        Do
            c = c + 1 'el contador sirve para la lista

            If mytable.EOF Then Exit Do
            'aqui es para abrir el file txt y llenarlo
            Filelibero1 = FreeFile
            Open FileName For Append As #Filelibero1
            conta = conta + 1
      
            conta_record = mytable.RecordCount

            'la cabecera
            If c = 1 Then
                myREG = myREG & "H"
                myREG = myREG & "|"
       
                'posicion 1 Serie/numero
                hastaCuanto = 4 - Trim(Len(mytable.Fields("SERIE")))
                Call E_llenar_zero(3 - Len(mytable.Fields("SERIE")), mytable.Fields("SERIE"), nuevoDato)
                myREG = myREG & "" & Trim(nuevoDato)
                myREG = myREG & "-"

                If Trim(mytable.Fields("NUMERO")) = "vacio" Then
                    myREG = myREG & "|"
                Else
                    hastaCuanto = 8 - Trim$(Len(mytable.Fields("NUMERO")))
                    myDato = mytable.Fields("NUMERO")
                    Call E_llenar_zero(hastaCuanto, myDato, nuevoDato)
                    myREG = myREG & nuevoDato
                    myREG = myREG & "|"

                End If

                'posicion 2 tipomoneda
                If Trim(mytable.Fields("MONEDA")) = Null Then
                    myREG = myREG & "|"
                Else
                    Call llenar_datos(3 - Trim(Len(mytable.Fields("MONEDA"))), mytable.Fields("MONEDA"), nuevoDato)
                    myREG = myREG & nuevoDato
                    myREG = myREG & "|"

                End If

                'posicion 3 Fecha de emisión
                If Trim(mytable.Fields("FECHA")) = 0 Then
                    myREG = myREG & "|"
                Else
                    Call llenar_datos(10 - Trim(Len(mytable.Fields("FECHA"))), mytable.Fields("FECHA"), nuevoDato)
                    myREG = myREG & nuevoDato
                    myREG = myREG & "|"

                End If

                'posicion 4 TipoDocIdentidadReceptor
                If Len(my_carga_busca_cliente(0).RUC) = 11 Then
                    myREG = myREG & "6" 'REG. UNICO DE CONTRIBUYENTES
                    myREG = myREG & "|"
                ElseIf Len(my_carga_busca_cliente(0).RUC) = 8 Then
                    myREG = myREG & "1" 'DOC. NACIONAL DE IDENTIDAD
                    myREG = myREG & "|"
                ElseIf Len(my_carga_busca_cliente(0).RUC) = 9 Then
                    myREG = myREG & "4" 'CARNET DE EXTRANJERIA
                    myREG = myREG & "|"
                Else
                    myREG = myREG & "0" 'OTROS
                    myREG = myREG & "|"

                End If
                
                'posicion 5 NumeroDocIdentidadReceptor
                If my_carga_busca_cliente(0).RUC = "" Then
                    myREG = myREG & "000000|"
                Else
                    Call E_llenar_datos(11 - Trim(Len(my_carga_busca_cliente(0).RUC)), Trim(my_carga_busca_cliente(0).RUC), nuevoDato) 'posicion 5
                    myREG = myREG & Trim(nuevoDato)
                    myREG = myREG & "|"

                End If

                'posicion 6 RazonSocialReceptor
                If my_carga_busca_cliente(0).nombre = "" Then
                    myREG = myREG & "PUBLICO General|"
                Else
                    myREG = myREG & Trim(my_carga_busca_cliente(0).nombre)
                    myREG = myREG & "|"

                End If
               
                '**Datos del cliente o receptor/DireccionReceptor
                'posicion 7
                myREG = myREG & Trim(my_carga_busca_cliente(0).direccion)
                myREG = myREG & "|"
   
                '**Datos del cliente o receptor/CorreoReceptor
                'posicion 8
                myREG = myREG & Trim(my_carga_busca_cliente(0).correo)
                myREG = myREG & "|" 'gion separador
        
                '***Totales factura Venta ***/TotalOPGravadas  aqui es el subtotal sin igv
                'posicion 9
                ' Testing Proyecto Facturacion Electronica
                Call E_llenar_datos(15 - Len(mytable.Fields("subtotal")), mytable.Fields("subtotal"), nuevoDato)
                nuevoDato = nuevoDato - (mytable.Fields("gravado"))
                myREG = myREG & Format(Trim(nuevoDato), "0.00")
                myREG = myREG & "|" 'gion separador 'es junto con bajo
                ' Testing Proyecto Facturacion Electronica

                '**Totales factura Venta ***/TotalOPNoGravadas
                'posicion 10
                If my_struc_datos_empresa(0).Toperacion = "I" Then
                    Call E_llenar_datos(15 - Len(mytable.Fields("gravado")), mytable.Fields("gravado"), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|" 'gion separador
                Else
                    myREG = myREG & "|" 'gion separador

                End If
  
                '**Totales factura Venta***/TotalOPExoneradas  para el caso de exportacion
                'posicion 11
                If my_struc_datos_empresa(0).Toperacion = "E" Then
                    Call E_llenar_datos(15 - Len(mytable.Fields("gravado")), mytable.Fields("gravado"), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|" 'gion separador 'es junto con bajo
                Else
                    myREG = myREG & "|" 'gion separador

                End If
       
                'AQUI FALTA LA POSICION 12
                myREG = myREG & "|"
    
                '**Totales factura Venta***/TotalOPExoneradas
                'posicion 13
                If Trim(mytable.Fields("tdescuento")) = "0" Then
                    myREG = myREG & "|"
                Else
                    myREG = myREG & Format(Trim(mytable.Fields("tdescuento")), "0.00")
                    myREG = myREG & "|" 'gion separador

                End If

                '***información Adicional/ TotalAnticipos
                'posicion 14
                If Trim(mytable.Fields("adetotal")) = "0" Then
                    myREG = myREG & "|"
                Else
                    Call E_llenar_datos(18 - Trim(Len(mytable.Fields("adetotal"))), mytable.Fields("adetotal"), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|" 'gion separador

                End If
 
                '**Totales de la Factura /Totales de la Boleta de Venta SumatoriaIGV
                'posicion 15
                Call E_llenar_datos(15 - Trim(Len(mytable.Fields("impuesto"))), mytable.Fields("impuesto"), nuevoDato)
                myREG = myREG & Format(Trim(nuevoDato), "0.00") 'aqui es de analizar
                myREG = myREG & "|" 'gion separador

                '**Totales de la Factura/Totales factura Venta***/TotalOPExoneradas
                'posicion 16
                If Trim(Len(Trim(mytable.Fields("tisc")))) = "0" Then
                    myREG = myREG & "|"
                Else '/SumatoriaISC 'posicion 16
                    Call E_llenar_datos(15 - Trim(Len(Trim(mytable.Fields("tisc")))), mytable.Fields("tisc"), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|" 'gion separador

                End If

                '**Totales de la Boleta de Venta  (Es identico a la Factura) ***/SumatoriaOtrosTributos
                'posicion 17
                If Trim(mytable.Fields("tdetra")) = "0" Then
                    myREG = myREG & "|"
                Else
                    Call E_llenar_datos(15 - Trim(Len(mytable.Fields("tdetra"))), Trim(mytable.Fields("tdetra")), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|" 'gion separador

                End If
    
                '**DescuentosGlobales
                ' posicion 18
                If Trim(mytable.Fields("tdescuento")) = "0" Then
                    myREG = myREG & "|"
                Else 'posicion 18
                    Call E_llenar_datos(15 - Trim(Len(mytable.Fields("tdescuento"))), mytable.Fields("tdescuento"), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00") '
                    myREG = myREG & "|" 'gion separador

                End If

                '**Sumatoria otros Cargos
                'POSICION 19
                If Trim("" & mytable.Fields("comision")) = "0" Then
                    myREG = myREG & "|"
                Else
                    Call E_llenar_datos(15 - Len("" & mytable.Fields("comision")), Trim("" & mytable.Fields("comision")), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|"

                End If

                '**ImporteTotalVenta
                'posicion 20
                Call E_llenar_datos(15 - Trim(Len(mytable.Fields("total"))), mytable.Fields("total"), nuevoDato)
                myREG = myREG & Format(Trim(Round(nuevoDato, 2)), "0.00")
                myREG = myREG & "|" 'gion separador

                '***Información Adicional/MontoEnLetras no es obligatorio que seria el posicion 21
                'posicion 21
                myREG = myREG & "|" 'gion separador
      
                '***Monto de la percepción
                'posicion 22
                If Trim("" & mytable.Fields("PERCEPCION")) = "0" Then
                    myREG = myREG & "|"
                Else
                    Call E_llenar_datos(15 - Trim(Len("" & mytable.Fields("PERCEPCION"))), "" & mytable.Fields("PERCEPCION"), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00") 'aqui es de analizar
                    myREG = myREG & "|"

                End If

                '***BaseImponiblePercepcion
                'posicion 23
                If Trim(mytable.Fields("percepcion")) = "0" Then
                    myREG = myREG & "|"
                Else ''posicion 23
                    Call E_llenar_datos(15 - Trim(Len(mytable.Fields("percepcion"))), "" & mytable.Fields("percepcion"), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|" 'gion separador

                End If

                '***Información Adicional - Percepciones
                'posicion 24
                If Trim(mytable.Fields("percepcion")) = "0" Then
                    myREG = myREG & "|"
                Else
                    Call E_llenar_datos(15 - Len(mytable.Fields("total") - Trim("" & mytable.Fields("percepcion"))), Val(mytable.Fields("total") - Trim("" & mytable.Fields("percepcion"))), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|" 'gion separador

                End If

                'posicion 25 PorcentajePercepcion
                myREG = myREG & "|"
                'CodigoTipoOperacion
                
                'posicion 26
                '''' 17/07/2018 Factura de Exportación
                myREG = myREG & my_tipooperacion
                myREG = myREG & "|"
                '''' 17/07/2018 Factura de Exportación

                '***CodigoLeyenda
                'posicion 27
                If my_struc_credito(0).anticipo = "" Then
                    myREG = myREG & "|" 'gion separador
                Else
                    myREG = myREG & "2000"
                    myREG = myREG & "|" 'gion separador

                End If

                '***Información Adicional/DescripcionLeyenda 'SerieNumeroAnticipo
                'posicion 28
                If Trim("" & mytable.Fields("percepcion")) = "0" Then
                    myREG = myREG & "|" 'gion separador
                Else
                    Call E_llenar_datos(100 - Len("COMPROBANTE DE PERCEPCION"), "COMPROBANTE DE PERCEPCION", nuevoDato)
                    myREG = myREG & Trim(nuevoDato) 'posicion 28
                    myREG = myREG & "|" 'gion separador

                End If

                '**SerieNumeroGuia
                ''posicion 29
                If Trim(mytable.Fields("serie1")) <> "" Then
                    Call E_llenar_datos(4 - Len(mytable.Fields("serie1")), mytable.Fields("serie1"), nuevoDato)
                    myREG = myREG & Trim(nuevoDato)
                    myREG = myREG & "-" 'gion separador

                End If

                '***Documentos de referencia/numero item 16 'posicion 29
                If mytable.Fields("numero1") <> "" Then
                    Call E_llenar_datos(8 - Len(mytable.Fields("numero1")), mytable.Fields("numero1"), nuevoDato)
                    myREG = myREG & Trim(nuevoDato) 'item 16
                    myREG = myREG & "|" 'gion separador
                Else
                    myREG = myREG & "|" 'gion separador

                End If

                '**Documentos de referencia/tipo
                'posicion 30  (Tipo de documento - Catálogo No. 12) SUNAT
                If Trim(mytable.Fields("tipo1")) <> "" Then
                    myREG = myREG & "99|"
                Else
                    myREG = myREG & "|"

                End If

                '**NumeroDocumentoRelacionad0
                'posicion 31
                If Trim(mytable.Fields("serie1")) <> "" Then
                    Call E_llenar_datos(4 - Len(mytable.Fields("serie1")), mytable.Fields("serie1"), nuevoDato)
                    myREG = myREG & Trim(nuevoDato) 'item 16
                    myREG = myREG & "-" 'gion separador

                End If

                '***TipoDocumentoRelacionado
                'posicion 31
                If mytable.Fields("numero1") <> "" Then
                    Call E_llenar_datos(8 - Len(mytable.Fields("numero1")), mytable.Fields("numero1"), nuevoDato)
                    myREG = myREG & Trim(nuevoDato) 'item 16
                    myREG = myREG & "|" 'gion separador
                Else
                    myREG = myREG & "|" 'gion separador

                End If

                '**TipoDocumentoGuia
                'Catálogo No. 12 Boleta de Venta/emitida por anticipos
                If Trim(mytable.Fields("tipo1")) <> "" Then
                    myREG = myREG & "99|"
                Else
                    myREG = myREG & "|"

                End If

                '***UbigeoPuntoLlegada
                'posicion 33
                Call E_llenar_datos(100 - Len(my_struc_ubigeo_Receptor(0).idubigeo), my_struc_ubigeo_Receptor(0).idubigeo, nuevoDato)
                myREG = myREG & Trim(nuevoDato)
                myREG = myREG & "|" 'gion separador
    
                '***DireccionPuntoLlegada
                'posicion 34
                Call E_llenar_datos(100 - Len(my_carga_busca_cliente(0).direccion), my_carga_busca_cliente(0).direccion, nuevoDato)
                myREG = myREG & Trim(nuevoDato) 'item 49
                myREG = myREG & "|" 'gion separador
    
                '***UrbanizacionPuntoLlegada
                'posicion 35
                myREG = myREG & "|" 'gion separador no hay en el sistema
   
                '***ProvinciaLugarEntrega
                'posicion 36
                Call E_llenar_datos(30 - Len(my_struc_ubigeo_Receptor(0).provincia), my_struc_ubigeo_Receptor(0).provincia, nuevoDato)
                myREG = myREG & Trim(nuevoDato) 'item 49
                myREG = myREG & "|" 'gion separador
  
                '**DepartamentoLugarEntrega
                'posicion 37
                Call E_llenar_datos(30 - Len(my_struc_ubigeo_Receptor(0).departamento), my_struc_ubigeo_Receptor(0).departamento, nuevoDato) 'item 49 /posicion 49
                myREG = myREG & Trim(nuevoDato) 'item 49
                myREG = myREG & "|" 'gion separador
    
                '**DistritoLugarEntrega item 49/posicion 50
                Call E_llenar_datos(30 - Len(my_struc_ubigeo_Receptor(0).distrito), my_struc_ubigeo_Receptor(0).distrito, nuevoDato) 'item 49/posicion 50
                myREG = myREG & Trim(nuevoDato) 'item 49
                myREG = myREG & "|" 'gion separador
    
                '**Dirección del cliente/CodigoPaisReceptor posicion 39
                myREG = myREG & "PE" ' muestra de b&H
                myREG = myREG & "|" 'gion separador 'posicion 39

                '***********************************************************
                'inicio 14/02/2018 pll forma de pago
                If "" & mytable.Fields("fpago") <> "Null" Then
                    If mytable.Fields("fpago") = "3" Then
                        Call E_llenar_datos(30 - Len("CREDITO"), "CREDITO", nuevoDato)
                        myREG = myREG & Trim(nuevoDato)
                    Else
                        Call E_llenar_datos(30 - Len("CONTADO"), "CONTADO", nuevoDato)
                        myREG = myREG & Trim(nuevoDato)

                    End If

                    'myREG = myREG & "|" 'gion separador
                End If

                'fin 14/02/2018 pll forma de pago
 
                ' Varios Locales FE 18/05/2018
                myDato = my_struc_datos_empresa(0).CodSede
                myREG = myREG & "|"

                If Len(myDato) > 0 Then
                    myREG = myREG & myDato

                End If

                ' Varios Locales FE 18/05/2018
   
                Print #Filelibero1, myREG
                'para el detalle del producto
 
                myREG = myREG & Chr(13)

            End If

            myREG = ""
            myREG = myREG & "I"
            myREG = myREG & "|"
     
            '**Datos del detalle
            
            'posicion 1 OrdenItem
            myDato = c

            If Len(myDato) = 1 Then
                myREG = myREG & "0" & myDato
                myREG = myREG & "|"
            Else
                myREG = myREG & myDato
                myREG = myREG & "|"

            End If

            'posicion 2 CodigoProductoItem
            If mytable.Fields("producto") = Null Then
                myREG = myREG & "|"
            Else
                Call llenar_datos(30 - Trim(Len(mytable.Fields("producto"))), mytable.Fields("producto"), nuevoDato)
                myREG = myREG & Trim(nuevoDato)
                myREG = myREG & "|"

            End If

            'posicion 3 DescripcionItem
            If mytable.Fields("descripcioP") = Null Then
                myREG = myREG & "|"
            Else
                Call E_llenar_datos(250 - Trim(Len(mytable.Fields("descripcioP"))), mytable.Fields("descripcioP"), nuevoDato)
                myREG = myREG & Trim(nuevoDato)
                myREG = myREG & "|"

            End If
            
            'posicion 4 UnidadMedidaItem
            If mytable.Fields("unidad") = Null Then
                myREG = myREG & "|"
            Else

                If mytable.Fields("pservicio") = "S" Then
                    myREG = myREG & "ZZ|"
                Else
                    myREG = myREG & "NIU|"

                End If

            End If
 
            'posicion 5 CantidadItem
            If mytable.Fields("cantidad") = Null Then
                myREG = myREG & "|"
            Else
                Call E_llenar_datos(23 - Trim(Len(mytable.Fields("cantidad"))), mytable.Fields("cantidad"), nuevoDato)
                myREG = myREG & Trim(nuevoDato)
                myREG = myREG & "|"

            End If

            'posicion 6 ValorVentaItem
            my_precioSinigv = mytable.Fields("precio") / (1 + mytable.Fields("igv") / 100)
            'my_dato = Round(my_precioSinigv, 2) * mytable.Fields("cantidad")
            my_dato = Round(my_precioSinigv, 2) * mytable.Fields("cantidad") - mytable.Fields("descuento")
            myREG = myREG & Format(Trim(my_dato), "0.00")
            myREG = myREG & "|"
    
            'posicion 7 ValorUnitarioSinIgv
            my_precioSinigv = 0
            my_precioSinigv = (mytable.Fields("precio") / (1 + mytable.Fields("igv") / 100))
            my_dato = Round(my_precioSinigv, 2)
            myREG = myREG & Format(Trim(my_dato), "0.00")
            myREG = myREG & "|"
  
            'posicion 8 PrecioUnitarioConIgv
            Call E_llenar_datos(15 - Len(mytable.Fields("precio")), mytable.Fields("precio"), nuevoDato)  'item 28
            myREG = myREG & Format(Trim(nuevoDato), "0.00")
            myREG = myREG & "|" 'gion separador 'es junto con bajo
  
            'posicion 9 DescuentoItem
            If Trim("" & mytable.Fields("descuento")) = "0" Then
                myREG = myREG & "0.00|"
            Else
                Call E_llenar_datos(15 - Trim(Len(mytable.Fields("descuento"))), Trim(mytable.Fields("descuento")), nuevoDato) 'item 28
                myREG = myREG & Format(Trim(nuevoDato), "0.00")
                myREG = myREG & "|"

            End If

            'posicion 10 CodTipoPrecioVtaUnitarioItem
            myREG = myREG & "01"
            myREG = myREG & "|"
            
            'posicion 11 ImporteIGVItem
            '''' 17/07/2018 Factura de Exportación
            Call E_llenar_datos(15 - Trim(Len(mytable.Fields("impuestod"))), mytable.Fields("impuestod"), nuevoDato) 'item 27
            myREG = myREG & Format(Trim(nuevoDato), "0.00")
            myREG = myREG & "|"
            '''' 17/07/2018 Factura de Exportación
        
            'posicion 12 CodigoAfectacionIGVItem
            '''' 17/07/2018 Factura de Exportación
            If mytable.Fields("igvd") = "0" Then
                If my_tipoigv = "40" Then
                    myREG = myREG & "40|"
                Else
                    myREG = myREG & "20|"

                End If

            Else
                myREG = myREG & "10|"

            End If

            '''' 17/07/2018 Factura de Exportación
 
            'posicion 13 MontoISCItem
            If Trim("" & mytable.Fields("dtisc")) = "" Or Trim("" & mytable.Fields("dtisc")) = Null Or Trim("" & mytable.Fields("dtisc")) = "0" Then
                myREG = myREG & "|"
            Else
                Call E_llenar_datos(15 - Trim(Len(mytable.Fields("dtisc"))), Trim(mytable.Fields("dtisc")), nuevoDato)
                myREG = myREG & Format(Trim(nuevoDato), "0.00")
                myREG = myREG & "|"

            End If
            
            'posicion 14 TipoSistemaISCItem
            myREG = myREG & "03|"
       
            'posicion 15 MontoReferencialUnitarioItem
            myREG = myREG & "0.00"
            myREG = myREG & "|"
    
            'posicion 16 CodigoTipoPrecioReferencial
            myREG = myREG & "|"
                    
            Print #Filelibero1, myREG
 
            ' Campos Adicionales FE 19/05/2018
 
            If c = conta_record Then
                myREG = myREG & Chr(13)
                myREG = ""
                myREG = myREG & "A"
                myREG = myREG & "|1||OBSERVACIÓN: "
                myREG = myREG & mytable.Fields("observa")
                Print #Filelibero1, myREG

            End If

            '  myREG = myREG & Chr(13)
            '   myREG = ""
            '   myREG = myREG & "A"
            '   myREG = myREG & "|1|100|Orden:ver-4"
            '   Print #Filelibero1, myREG
            '
            '    myREG = myREG & Chr(13)
            '   myREG = ""
            '   myREG = myREG & "A"
            '   myREG = myREG & "|1|100|xxxx"
            '   Print #Filelibero1, myREG
            ' Campos Adicionales FE 19/05/2018

            Close #Filelibero1
            mytable.MoveNext
        Loop

    End If
   
    mytable.Close
    'aqui copia el archivo txt

    origen = FileName

    destino = "D:\ce_Input\" & file

    ' Testing Proyecto Facturacion Electronica 28/02/2018
    'FileCopy origen, destino
    Dim fso As New Scripting.FileSystemObject

    fso.MoveFile origen, destino
    ' Testing Proyecto Facturacion Electronica  28/02/2018

End Function

Function estrae_nota_credito(my_ruc As String, _
                             local1 As String, _
                             bxtipo As String, _
                             bxserie As String, _
                             bxnumero As String, _
                             my_idubigeo As String, _
                             acu As String, _
                             my_struc_datos_empresa() As struc_datos_empresa, _
                             my_struc_ubigeo_Receptor() As struc_ubigeo_Receptor, _
                             my_carga_busca_cliente() As struc_busca_cliente, _
                             my_struc_credito() As struc_credito, _
                             paso As Boolean, _
                             my_struc_ubigeo_Emisor() As struc_ubigeo_Emisor, _
                             file As String, _
                             my_tipo As String)

    Dim mysql           As String

    Dim mytable         As New ADODB.Recordset

    Dim hastaCuanto     As Integer

    Dim nuevoDato       As String

    Dim myDato          As String

    Dim my_precioSinigv As Currency

    Dim myREG           As String

    Dim xnumero         As String

    Dim xserie          As String

    mysql = ""
    mysql = "SELECT d.impuesto*-1 as impuestod, d.igv as igvd, f.SERIE," & Chr$(10)
    mysql = mysql & "f.NUMERO, " & Chr$(10) 'serie y número correlativo item1
    mysql = mysql & "CONVERT(VARCHAR(10),f.FECHA, 120) AS FECHA ," & Chr$(10) 'Fecha de emisión item 3
    mysql = mysql & "CONVERT(VARCHAR(10),f.FECHAE, 120) AS FECHAE ," & Chr$(10) 'Fechavencimiento item 4
    mysql = mysql & "f.tipo," & Chr$(10) 'TipoDocumento item 3
    mysql = mysql & "Case f.moneda" & Chr$(10) ' TipoMoneda item 4
    mysql = mysql & "WHEN 'S' THEN 'PEN'" & Chr$(10) ' TipoMoneda item 4
    mysql = mysql & "WHEN 'D' THEN 'USD'" & Chr$(10) ' TipoMoneda item 4
    mysql = mysql & "END AS MONEDA," & Chr$(10) ' TipoMoneda item 4
    '****aqui todo con parametrizacion va al txt
    'Documentos de referencia
    mysql = mysql & "f.adetotal," & Chr$(10) 'serie item 15 'adelanto total
    mysql = mysql & "f.serie1," & Chr$(10) 'serie item 16
    mysql = mysql & "f.numero1," & Chr$(10) 'numero item 16
    mysql = mysql & "f.tipo1," & Chr$(10) 'tipo item 16
    mysql = mysql & "f.tdetra," & Chr$(10) 'tipo item 18
    mysql = mysql & "f.partida," & Chr$(10)
    mysql = mysql & "f.destino," & Chr$(10)
    mysql = mysql & "d.unidad," & Chr$(10)  'UnidadMedidaItem
    mysql = mysql & "d.cantidad *-1 as cantidad," & Chr$(10)  'CantidadItem
    mysql = mysql & "d.producto," & Chr$(10) 'Producto
    mysql = mysql & "d.descripcio as descripcioP," & Chr$(10) 'DescripcionItem producto
    mysql = mysql & "isnull(p.costoinisigv,0)*-1 as costoinisigv," & Chr$(10) 'ValorUnitarioSinIgv
    mysql = mysql & "d.precio," & Chr$(10) 'PrecioUnitarioConIgv
    'CodTipoPrecioVtaUnitarioItem 01 precio unitario, 02 valor referencial txt
    mysql = mysql & "f.subtotal*-1 as subtotal," & Chr$(10) 'ImporteIGVItem
    'CodigoAfectacionIGVItem **Afectación al IGV - Catálogo No. 07 txt
    mysql = mysql & "f.tisc," & Chr$(10) 'MontoISCItem
    mysql = mysql & "d.tisc as dtisc," & Chr$(10) 'MontoISCItem
    mysql = mysql & "isnull(p.costou,0) as costou," & Chr$(10) 'ValorVentaItem
    mysql = mysql & "f.impuesto*-1 as impuesto," & Chr$(10) 'ValorVentaItem
    mysql = mysql & "p.igv," & Chr$(10) 'posicion 8
    mysql = mysql & "f.descuento*-1 as tdescuento, " & Chr$(10) 'DescuentoItem
    mysql = mysql & "d.descuento as descuento, " & Chr$(10) 'DescuentoItem
    mysql = mysql & "(d.cantidad * d.precio)*-1 as ValorVenta," & Chr$(10) 'ValorVentaItem" y la deduccion de los descuentos
    'Totales de la Boleta de Venta  (Es identico a la Factura) ***
    mysql = mysql & "f.total*-1 as total," & Chr$(10) 'TotalOPGravadas item 30
    mysql = mysql & "isnull(f.fpago,0) as fpago," & Chr$(10)
    mysql = mysql & "d.subtotal*-1 as subtotal," & Chr$(10) 'SumatoriaIGV item 35 **se calcula text
    mysql = mysql & "d.comision as comision," & Chr$(10)
    '**Documentos de referencia/TipoDocumentoRelacionado
    'mysql = mysql & "f.acu," & Chr$(10) 'TipoDocumentoRelacionado posicion 41
    '**Posicion 10/CodTipoPrecioVtaUnitarioItem 01 incluye igv 02 operaciones no honerosas
    mysql = mysql & "f.impuesto*-1  as impuesto," & Chr$(10)
    mysql = mysql & "f.subtotal*-1 as subtotal," & Chr$(10)
    mysql = mysql & "f.neto*-1 as neto," & Chr$(10)
    mysql = mysql & "f.gravado*-1 as gravado," & Chr$(10)
 
    '***Información adicional - Percepciones
    mysql = mysql & "f.percepcion," & Chr$(10) 'BaseImponiblePercepcion 41
    mysql = mysql & "d.tpercepcio," & Chr$(10) ' MontoPercepcion 41 *aqui calcular txt
    '***Información Adicional - Factura Guía /Marca y placa de Vehiculo item 52

    ' Testing Proyecto Facturacion Electronica
    mysql = mysql & "f.TIPONCD,f.observa, p.servicio as pservicio" & Chr$(10) ' MontoPercepcion 41 *aqui calcular txt
    ' Testing Proyecto Facturacion Electronica

    mysql = mysql & "from factura f," & Chr$(10)
    mysql = mysql & "detalle d," & Chr$(10)
    mysql = mysql & "producto p" & Chr$(10)
    mysql = mysql & "where f.local='" & "" & local1 & "' " & Chr$(10)
    mysql = mysql & "and f.serie ='" & "" & bxserie & "' " & Chr$(10)
    mysql = mysql & "and f.numero='" & "" & bxnumero & "' " & Chr$(10)
    mysql = mysql & "and f.acu='" & "" & acu & "' " & Chr$(10)
    mysql = mysql & "and f.tipo='" & "" & my_tipo & "' " & Chr$(10)
    mysql = mysql & "and f.SERIE = d.serie" & Chr$(10)
    mysql = mysql & "and f.NUMERO = d.NUMERO" & Chr$(10)
    mysql = mysql & "and p.producto = d.PRODUCTO" & Chr$(10)
    mysql = mysql & "and f.tipo = d.tipo" & Chr$(10)
     
    mytable.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytable.RecordCount > 0 Then
        'para el numero
        hastaCuanto = 8 - Len(mytable.Fields("NUMERO"))
        myDato = mytable.Fields("NUMERO")
        Call E_llenar_zero(hastaCuanto, myDato, xnumero)
        'PARA LA SERIE
        Call E_llenar_zero(3 - Len(mytable.Fields("SERIE")), mytable.Fields("SERIE"), xserie)

        If bxtipo = "2" Then
            FileName = "D:\ce_output\CREA\" & my_ruc & "_07" & "_" & xserie & "-" & xnumero & ".INPUT.TXT"
            file = my_ruc & "_07" & "_" & xserie & "-" & xnumero & ".INPUT.TXT"
        ElseIf bxtipo = "1" Then
            FileName = "D:\ce_output\CREA\" & my_ruc & "_07" & "_" & xserie & "-" & xnumero & ".INPUT.TXT"
            file = my_ruc & "_07" & "_" & xserie & "-" & xnumero & ".INPUT.TXT"

        End If
 
        Do
            c = c + 1 'el contador sirve para la lista

            If mytable.EOF Then Exit Do
            'aqui es para abrir el file txt y llenarlo
            Filelibero1 = FreeFile
            Open FileName For Append As #Filelibero1
            conta = conta + 1
      
            conta_record = mytable.RecordCount

            'myREG = ""
            'la cabecera
            If c = 1 Then
                myREG = myREG & "H"
                myREG = myREG & "|"
                'posicion 1
    
                If bxtipo = "2" Then
                    hastaCuanto = 4 - Trim(Len(mytable.Fields("SERIE")))
                    Call E_llenar_zero(3 - Len(mytable.Fields("SERIE")), mytable.Fields("SERIE"), nuevoDato)
                    myREG = myREG & "" & Trim(nuevoDato)
                    myREG = myREG & "-"
                Else
                    hastaCuanto = 4 - Trim(Len(mytable.Fields("SERIE")))
                    Call E_llenar_zero(3 - Len(mytable.Fields("SERIE")), mytable.Fields("SERIE"), nuevoDato)
                    myREG = myREG & "" & Trim(nuevoDato)
                    myREG = myREG & "-"

                End If
    
                'posicion 2
                If Trim(mytable.Fields("NUMERO")) = "vacio" Then
                    myREG = myREG & "|" 'gion separador
                Else
                    hastaCuanto = 8 - Trim$(Len(mytable.Fields("NUMERO")))
                    myDato = mytable.Fields("NUMERO")
                    Call E_llenar_zero(hastaCuanto, myDato, nuevoDato)
                    myREG = myREG & nuevoDato
                    myREG = myREG & "|" 'gion separador

                End If
   
                'posicion 2
                If Trim(mytable.Fields("FECHAE")) = 0 Then
                    Call E_llenar_datos(9, "0", nuevoDato)
                    myREG = myREG & nuevoDato
                    myREG = myREG & "|" 'gion separador
                Else
                    hastaCuanto = 10 - Len(mytable.Fields("FECHAE"))
                    myDato = mytable.Fields("FECHAE")
                    Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'NUMERO 3
                    myREG = myREG & nuevoDato
                    myREG = myREG & "|" 'gion separador

                End If

                '**SerieNumeroModifica
                'posicion 3
   
                ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
                If bxtipo = "2" Then
                    hastaCuanto = 4 - Trim(Len(mytable.Fields("SERIE1")))
                    myDato = mytable.Fields("SERIE1")
                    Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                    myREG = myREG & "" & Trim(nuevoDato)
                    myREG = myREG & "-"
                Else
                    hastaCuanto = 4 - Trim(Len(mytable.Fields("SERIE1")))
                    Call E_llenar_zero(3 - Len(mytable.Fields("SERIE1")), mytable.Fields("SERIE"), nuevoDato)
                    myREG = myREG & "" & Trim(nuevoDato)
                    myREG = myREG & "-"

                End If
      
                If Trim(mytable.Fields("NUMERO1")) = "vacio" Then
                    myREG = myREG & "|"
                Else
                    hastaCuanto = 8 - Trim$(Len(mytable.Fields("NUMERO1"))) '**en la tabla
                    myDato = mytable.Fields("NUMERO1")
                    Call E_llenar_zero(hastaCuanto, myDato, nuevoDato) 'NUMERO 3
                    myREG = myREG & nuevoDato
                    myREG = myREG & "|" 'gion separador

                End If

                ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
 
                'posicion 4
                If bxtipo = "2" Then
                    myREG = myREG & "01"
                    myREG = myREG & "|" 'gion separador
                ElseIf bxtipo = "1" Then
                    myREG = myREG & "03"
                    myREG = myREG & "|" 'gion separador

                End If

                'TipoNotaCredito
                'posicion 5
  
                If mytable.Fields("TIPONCD") = Null Then 'posicion 6
                    myREG = myREG & "|"
                Else
                    myREG = myREG & mytable.Fields("TIPONCD")
                    myREG = myREG & "|" 'gion separador

                End If

                'MotivoSustento
                'posicion 6
                If mytable.Fields("observa") = Null Then
                    myREG = myREG & "|"
                Else
                    Call llenar_datos(250 - Trim(Len(mytable.Fields("observa"))), mytable.Fields("observa"), nuevoDato) 'tipoclie 4
                    myREG = myREG & Trim(nuevoDato)
                    myREG = myREG & "|" 'gion separador

                End If

                'posicion 7
                If Trim(mytable.Fields("MONEDA")) = Null Then
                    myREG = myREG & "|"
                Else
                    Call llenar_datos(3 - Len(mytable.Fields("MONEDA")), mytable.Fields("MONEDA"), nuevoDato)  'MONEDA 9
                    myREG = myREG & nuevoDato
                    myREG = myREG & "|" 'gion separador

                End If

                'Datos del Emisor **NumeroDocIdentidadReceptor 'Ruc
                'posicion 8
      
                If my_carga_busca_cliente(0).RUC = "" Then
                    myREG = myREG & "000000|"
                Else
                    Call E_llenar_datos(11 - Trim(Len(my_carga_busca_cliente(0).RUC)), Trim(my_carga_busca_cliente(0).RUC), nuevoDato) 'posicion 5
                    myREG = myREG & Trim(nuevoDato) 'item 12
                    myREG = myREG & "|" 'gion separador

                End If

                If Len(my_carga_busca_cliente(0).RUC) = 11 Then
                    myREG = myREG & "6" 'REG. UNICO DE CONTRIBUYENTES
                    myREG = myREG & "|"
                ElseIf Len(my_carga_busca_cliente(0).RUC) = 8 Then
                    myREG = myREG & "1" 'DOC. NACIONAL DE IDENTIDAD
                    myREG = myREG & "|"
                ElseIf Len(my_carga_busca_cliente(0).RUC) = 9 Then
                    myREG = myREG & "4" 'CARNET DE EXTRANJERIA
                    myREG = myREG & "|"
                Else
                    myREG = myREG & "0" 'OTROS
                    myREG = myREG & "|"

                End If

                '**Datos del cliente o receptor/RazonSocialReceptor
                'posicion 6
                If my_carga_busca_cliente(0).nombre = "" Then
                    myREG = myREG & "PUBLICO General|"
                Else
                    myREG = myREG & Trim(my_carga_busca_cliente(0).nombre)
                    myREG = myREG & "|"

                End If

                myREG = myREG & Trim(my_struc_ubigeo_Receptor(0).idubigeo)
                myREG = myREG & "|" 'gion separador

                '**Datos del cliente o receptor/DireccionReceptor
                'posicion 7
                myREG = myREG & Trim(my_carga_busca_cliente(0).direccion)
                myREG = myREG & "|" 'gion separador
   
                myREG = myREG & ""
                myREG = myREG & "|" 'gion separador
 
                myREG = myREG & ""
                myREG = myREG & "|" 'gion separador
   
                myREG = myREG & Trim(my_carga_busca_cliente(0).dpto)
                myREG = myREG & "|" 'gion separador
   
                myREG = myREG & ""
                myREG = myREG & "|" 'gion separador
     
                '**Dirección del cliente/CodigoPaisReceptor
                'posicion 17
                myREG = myREG & "PE" ' muestra de b&H
                myREG = myREG & "|" 'gion separador

                '**Datos del cliente o receptor/CorreoReceptor
                'posicion 18
                Call E_llenar_datos(100 - Len(my_carga_busca_cliente(0).correo), my_carga_busca_cliente(0).correo, nuevoDato)
                myREG = myREG & Trim(nuevoDato) 'item 15
                myREG = myREG & "|" 'gion separador
                ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

                ' Testing Proyecto Facturacion Electronica
                '**TotalOPGravadas
                'posicion 19
                Call E_llenar_datos(15 - Len(mytable.Fields("subtotal")), mytable.Fields("subtotal"), nuevoDato)
                nuevoDato = nuevoDato - mytable.Fields("gravado")
                myREG = myREG & Format(Trim(nuevoDato), "0.00")
                myREG = myREG & "|" 'gion separador 'es junto con bajo
                ' Testing Proyecto Facturacion Electronica

                '**TotalOPNoGravadas
                'posicion 20
                If my_struc_datos_empresa(0).Toperacion = "I" Then
                    Call E_llenar_datos(15 - Len(mytable.Fields("gravado")), mytable.Fields("gravado"), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|" 'gion separador
                Else
                    myREG = myREG & "|" 'gion separador

                End If

                '**TotalOPExonerado
                'posicion 21
                If my_struc_datos_empresa(0).Toperacion = "E" Then
                    Call E_llenar_datos(15 - Len(mytable.Fields("gravado")), mytable.Fields("gravado"), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|" 'gion separador 'es junto con bajo
                Else
                    myREG = myREG & "|" 'gion separador

                End If
 
                '**Totales factura Venta***/TotalOPExoneradas
                'posicion 22
                'TotalDescuentos
                If Trim("" & mytable.Fields("tdescuento")) = "0" Then
                    myREG = myREG & "|"
                Else
                    myREG = myREG & Format(Trim("" & mytable.Fields("tdescuento")), "0.00")
                    myREG = myREG & "|" 'gion separador

                End If

                '**SumatoriaIGV
                'posicion 23
                If "" & mytable.Fields("tipo") = "7" Then
                    myREG = myREG & "|" 'gion separador
                Else
                    Call E_llenar_datos(15 - Trim(Len(mytable.Fields("impuesto"))), mytable.Fields("impuesto"), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|" 'gion separador

                End If
  
                '**SumatoriaISC
                'posicion 24
                If "" & mytable.Fields("tisc") = "0" Then
                    myREG = myREG & "|"
                Else
                    Call E_llenar_datos(15 - Trim(Len("" & mytable.Fields("tisc"))), Trim("" & mytable.Fields("tisc")), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|" 'gion separador

                End If
  
                '**SumatoriaOtrosTributos
                'posicion 25
                If Trim(mytable.Fields("tdetra")) = "0" Then
                    myREG = myREG & "|"
                Else
                    Call E_llenar_datos(15 - Trim(Len(mytable.Fields("tdetra"))), Trim(mytable.Fields("tdetra")), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|" 'gion separador

                End If
    
                '**SumatoriaOtrosCargos
                'posicion 26
                If Trim("" & mytable.Fields("comision")) = "0" Then
                    myREG = myREG & "|"
                Else
                    Call E_llenar_datos(15 - Trim(Len("" & mytable.Fields("comision"))), "" & mytable.Fields("comision"), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|" 'gion separador

                End If

                '**ImporteTotal
                'posicion 27
                Call E_llenar_datos(15 - Trim(Len(mytable.Fields("total"))), mytable.Fields("total"), nuevoDato)
                myREG = myREG & Format(Trim(Round(nuevoDato, 2)), "0.00")
                myREG = myREG & "|"

                '***TotalAnticipos
                'posicion 28
                If Trim(mytable.Fields("adetotal")) = "0" Then
                    myREG = myREG & "|"
                Else
                    Call E_llenar_datos(18 - Trim(Len(mytable.Fields("adetotal"))), mytable.Fields("adetotal"), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|" 'gion separador

                End If

                ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
                myREG = myREG & "|"
                myREG = myREG & "|"
                myREG = myREG & "|"

                Print #Filelibero1, myREG
                'para el detalle del producto
 
                myREG = myREG & Chr(13)

            End If

            myREG = ""
            myREG = myREG & "I"
            myREG = myREG & "|"
            
            '**Datos del detalle
            'posicion 1 OrdenItem
            myDato = c

            If Len(myDato) = 1 Then
                myREG = myREG & "0" & myDato
                myREG = myREG & "|"
            Else
                myREG = myREG & myDato
                myREG = myREG & "|"

            End If

            'posicion 2 UnidadMedidaItem
            If mytable.Fields("unidad") = Null Then
                myREG = myREG & "|"
            Else

                If mytable.Fields("pservicio") = "S" Then
                    myREG = myREG & "ZZ|"
                Else
                    myREG = myREG & "NIU|"

                End If

            End If
            
            'posicion 3 CantidadItem
            If mytable.Fields("cantidad") = Null Then
                myREG = myREG & "|"
            Else
                Call E_llenar_datos(23 - Trim(Len(mytable.Fields("cantidad"))), mytable.Fields("cantidad"), nuevoDato)
                myREG = myREG & Trim(nuevoDato)
                myREG = myREG & "|"

            End If
            
            'posicion 4 CodigoProductoItem
            If mytable.Fields("producto") = Null Then
                myREG = myREG & "|"
            Else
                Call llenar_datos(30 - Trim(Len(mytable.Fields("producto"))), mytable.Fields("producto"), nuevoDato)
                myREG = myREG & Trim(nuevoDato)
                myREG = myREG & "|"

            End If

            'posicion 5 DescripcionItem
            If mytable.Fields("descripcioP") = Null Then
                myREG = myREG & "|"
            Else
                Call E_llenar_datos(250 - Trim(Len(mytable.Fields("descripcioP"))), mytable.Fields("descripcioP"), nuevoDato)
                myREG = myREG & Trim(nuevoDato)
                myREG = myREG & "|"

            End If
    
            'posicion 6 ValorUnitarioSinIgv
            my_precioSinigv = 0
            my_precioSinigv = (mytable.Fields("precio") / (1 + mytable.Fields("igv") / 100))
            my_dato = Round(my_precioSinigv, 2)
            myREG = myREG & Format(Trim(my_dato), "0.00")
            myREG = myREG & "|"
  
            'posicion 7 PrecioUnitarioConIgv
            Call E_llenar_datos(15 - Len(mytable.Fields("precio")), mytable.Fields("precio"), nuevoDato)  'item 28
            myREG = myREG & Format(Trim(nuevoDato), "0.00")
            myREG = myREG & "|" 'gion separador 'es junto con bajo
  
            'posicion 8 CodTipoPrecioVtaUnitarioItem
            myREG = myREG & "01"
            myREG = myREG & "|"
            
            'posicion 9 ImporteIGVItem
            '''' 17/07/2018 Factura de Exportación
            Call E_llenar_datos(15 - Trim(Len(mytable.Fields("impuestod"))), mytable.Fields("impuestod"), nuevoDato) 'item 27
            myREG = myREG & Format(Trim(nuevoDato), "0.00")
            myREG = myREG & "|"
            '''' 17/07/2018 Factura de Exportación
            
            'posicion 10 CodigoAfectacionIGVItem
            '''' 17/07/2018 Factura de Exportación
            If mytable.Fields("igvd") = "0" Then
                If my_tipoigv = "40" Then
                    myREG = myREG & "40|"
                Else
                    myREG = myREG & "20|"

                End If

            Else
                myREG = myREG & "10|"

            End If

            'posicion 11 MontoISCItem
            If Trim("" & mytable.Fields("dtisc")) = "" Or Trim("" & mytable.Fields("dtisc")) = Null Or Trim("" & mytable.Fields("dtisc")) = "0" Then
                myREG = myREG & "|"
            Else
                Call E_llenar_datos(15 - Trim(Len(mytable.Fields("dtisc"))), Trim(mytable.Fields("dtisc")), nuevoDato)
                myREG = myREG & Format(Trim(nuevoDato), "0.00")
                myREG = myREG & "|"

            End If

            'posicion 12 TipoSistemaISCItem
            myREG = myREG & "03|"

            'posicion 13 ValorVentaItem
            my_precioSinigv = mytable.Fields("precio") / (1 + mytable.Fields("igv") / 100)
            'my_dato = Round(my_precioSinigv, 2) * mytable.Fields("cantidad")
            my_dato = Round(my_precioSinigv, 2) * mytable.Fields("cantidad") - mytable.Fields("descuento")
            myREG = myREG & Format(Trim(my_dato), "0.00")
    
            'posicion 14 MontoReferenciaItem
            myREG = myREG & "|"

            'posicion 15 CodigoTipoPrecioReferencial
            myREG = myREG & "|"
            Print #Filelibero1, myREG

            ' Campos Adicionales FE 19/05/2018
            If c = conta_record Then
                myREG = myREG & Chr(13)
                myREG = ""
                myREG = myREG & "A"
                myREG = myREG & "|1||OBSERVACIÓN: "
                myREG = myREG & mytable.Fields("observa")
                Print #Filelibero1, myREG

            End If

            '  myREG = myREG & Chr(13)
            '   myREG = ""
            '   myREG = myREG & "A"
            '   myREG = myREG & "|1|100|Orden:ver-4"
            '   Print #Filelibero1, myREG
            '
            '    myREG = myREG & Chr(13)
            '   myREG = ""
            '   myREG = myREG & "A"
            '   myREG = myREG & "|1|100|xxxx"
            '   Print #Filelibero1, myREG
            ' Campos Adicionales FE 19/05/2018
   
            Close #Filelibero1
            mytable.MoveNext
        Loop
        Close #Filelibero1

    End If

    mytable.Close

    'aqui copia el archivo txt

    origen = FileName

    destino = "D:\ce_Input\" & file

    ' Testing Proyecto Facturacion Electronica 28/02/2018
    'FileCopy origen, destino
    Dim fso As New Scripting.FileSystemObject

    fso.MoveFile origen, destino
    ' Testing Proyecto Facturacion Electronica  28/02/2018

End Function

Function b_ubigeo_receptor(my_codcliente As String, _
                           salida As Boolean, _
                           my_struc_ubigeo_Receptor() As struc_ubigeo_Receptor)

    Dim mysql    As String

    Dim mytable  As New ADODB.Recordset

    Dim mytablex As New ADODB.Recordset

    ReDim my_struc_ubigeo_Receptor(0)

    mysql = ""
    mysql = "select isnull(DPTO,'vacio') as dpto," & Chr$(10)
    mysql = mysql & "isnull(DISTRITO,'vacio') as DISTRITO," & Chr$(10)
    mysql = mysql & "isnull(direccion,'vacio') as  direccion" & Chr$(10)
    mysql = mysql & "From clientes" & Chr$(10)
    mysql = mysql & "where codigo= '" & "" & Trim(my_codcliente) & "' " & Chr$(10)

    mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
 
    If mytable.EOF Then
        salida = False
        Exit Function
    Else
        salida = True

        If mytable.Fields("DPTO") = "vacio" Or mytable.Fields("DPTO") = "" Then
            my_dpto = ""
        Else
            my_dpto = mytable.Fields("DPTO")

        End If

        If mytable.Fields("DISTRITO") = "vacio" Or mytable.Fields("DISTRITO") = "" Then
            my_distrito = ""
        Else
            my_distrito = mytable.Fields("DISTRITO")

        End If

        If mytable.Fields("direccion") = "vacio" Or mytable.Fields("direccion") = "" Then
            my_direccion = ""
        Else
            my_direccion = mytable.Fields("direccion")

        End If

    End If
  
    mysql = ""
    mysql = "select idubigeo,provincia,departamento,distrito" & Chr$(10)
    mysql = mysql & "From ubigeo" & Chr$(10)
    mysql = mysql & "where Departamento='" & "" & my_dpto & "' " & Chr$(10)
    mysql = mysql & "and distrito='" & "" & my_distrito & "' " & Chr$(10)

    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic
      
    ''''' 17/07/2018 Factura de Exportación
    'If mytablex.EOF Then
    ' salida = False
    'Exit Function
    ''''' 17/07/2018 Factura de Exportación

    salida = True

    'mytablex.MoveFirst
    ' k = 0
    'Do Until mytablex.EOF
    If k > 0 Then
        ReDim Preserve my_struc_ubigeo_Receptor(UBound(my_struc_ubigeo_Receptor) + 1)

    End If
 
    my_struc_ubigeo_Receptor(UBound(my_struc_ubigeo_Receptor)).direccion = mytable.Fields("direccion")

    mytablex.Close
    mytable.Close

End Function

Public Function E_llenar_datos(hastaCuanto As Integer, _
                               myDato As String, _
                               nuevoDato As String)

    For I = 1 To hastaCuanto
        myDato = myDato & " "
    Next
    nuevoDato = myDato

End Function

Public Function E_llenar_zero(hastaCuanto As Integer, _
                              myDato As String, _
                              nuevoDato As String)

    For I = 1 To hastaCuanto
        myDato = "0" & myDato
    Next
    nuevoDato = myDato

End Function

Function b_credito(xserie As String, _
                   xnumero As String, _
                   my_credito As Boolean, _
                   my_struc_credito() As struc_credito)

    Dim mysql   As String

    Dim mytable As New ADODB.Recordset

    ReDim my_struc_credito(0)

    mysql = ""
    mysql = "select anticipo,fpago," & Chr$(10)
    mysql = mysql & "serie,numero," & Chr$(10)
    mysql = mysql & "total "
    mysql = mysql & "From cuentac" & Chr$(10)
    mysql = mysql & "where serie='" & "" & xserie & "' " & Chr$(10)
    mysql = mysql & "and numero='" & "" & xnumero & "' " & Chr$(10)
  
    mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
 
    If mytable.EOF Then
        my_credito = False
        Exit Function
    Else
        my_credito = True
        mytable.MoveFirst
        k = 0

        Do Until mytable.EOF

            If k > 0 Then
                ReDim Preserve my_struc_credito(UBound(my_struc_credito) + 1)

            End If
     
            If mytable.Fields("anticipo") <> "" Then
                my_struc_credito(UBound(my_struc_credito)).anticipo = mytable.Fields("anticipo")
            Else
                my_struc_credito(UBound(my_struc_credito)).anticipo = ""

            End If
      
            If mytable.Fields("fpago") <> "" Then
                my_struc_credito(UBound(my_struc_credito)).fpago = mytable.Fields("fpago")
            Else
                my_struc_credito(UBound(my_struc_credito)).fpago = ""

            End If
      
            If mytable.Fields("serie") <> "" Then
                my_struc_credito(UBound(my_struc_credito)).serie = mytable.Fields("serie")
            Else
                my_struc_credito(UBound(my_struc_credito)).serie = ""

            End If
    
            If mytable.Fields("numero") <> "" Then
                my_struc_credito(UBound(my_struc_credito)).Numero = mytable.Fields("numero")
            Else
                my_struc_credito(UBound(my_struc_credito)).Numero = ""

            End If
    
            If mytable.Fields("total") <> "" Then
                my_struc_credito(UBound(my_struc_credito)).total = mytable.Fields("total")
            Else
                my_struc_credito(UBound(my_struc_credito)).total = ""

            End If
   
            k = k + 1
            mytable.MoveNext
        Loop

    End If

    mytable.Close

End Function

Function b_ubigeo_emisor(my_ruc As String, _
                         salida As Boolean, _
                         my_struc_ubigeo_Emisor() As struc_ubigeo_Emisor)

    Dim mysql    As String

    Dim mytable  As New ADODB.Recordset

    Dim mytablex As New ADODB.Recordset

    ReDim my_struc_ubigeo_Emisor(0)

    mysql = ""
    mysql = "select DPTO,DISTRITO,direccion" & Chr$(10)
    mysql = mysql & "From empresa" & Chr$(10)
    mysql = mysql & "where codigo1='" & "" & my_ruc & "' " & Chr$(10)

    mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
 
    If mytable.EOF Then
        salida = False
        Exit Function
    Else
        salida = True
        my_dpto = mytable.Fields("DPTO")
        my_distrito = mytable.Fields("DISTRITO")

    End If
  
    mysql = ""
    mysql = "select idubigeo,provincia,departamento,distrito" & Chr$(10)
    mysql = mysql & "From ubigeo" & Chr$(10)
    mysql = mysql & "where Departamento='" & "" & my_dpto & "' " & Chr$(10)
    mysql = mysql & "and distrito='" & "" & my_distrito & "' " & Chr$(10)

    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF Then
        salida = False
        Exit Function
    Else
        salida = True
        mytablex.MoveFirst
        k = 0

        Do Until mytablex.EOF

            If k > 0 Then
                ReDim Preserve my_struc_ubigeo_Emisor(UBound(my_struc_ubigeo_Emisor) + 1)

            End If
     
            If mytablex.Fields("idubigeo") <> "" Then
                my_struc_ubigeo_Emisor(UBound(my_struc_ubigeo_Emisor)).idubigeo = mytablex.Fields("idubigeo")
            Else
                my_struc_ubigeo_Emisor(UBound(my_struc_ubigeo_Emisor)).idubigeo = ""

            End If
      
            If mytablex.Fields("provincia") <> "" Then
                my_struc_ubigeo_Emisor(UBound(my_struc_ubigeo_Emisor)).provincia = mytablex.Fields("provincia")
            Else
                my_struc_ubigeo_Emisor(UBound(my_struc_ubigeo_Emisor)).provincia = ""

            End If
      
            If mytablex.Fields("departamento") <> "" Then
                my_struc_ubigeo_Emisor(UBound(my_struc_ubigeo_Emisor)).departamento = mytablex.Fields("departamento")
            Else
                my_struc_ubigeo_Emisor(UBound(my_struc_ubigeo_Emisor)).departamento = ""

            End If
    
            If mytablex.Fields("distrito") <> "" Then
                my_struc_ubigeo_Emisor(UBound(my_struc_ubigeo_Emisor)).distrito = mytablex.Fields("distrito")
            Else
                my_struc_ubigeo_Emisor(UBound(my_struc_ubigeo_Emisor)).distrito = ""

            End If

            'direccion
            If mytable.Fields("direccion") <> "" Then
                my_struc_ubigeo_Emisor(UBound(my_struc_ubigeo_Emisor)).direccion = mytable.Fields("direccion")
            Else
                my_struc_ubigeo_Emisor(UBound(my_struc_ubigeo_Emisor)).direccion = ""

            End If
   
            k = k + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
    mytable.Close
 
End Function

Public Function read_save_electronico(input_file As String, _
                                      my_local As String, _
                                      my_serie As String, _
                                      my_numero As String, _
                                      my_tipo As String, _
                                      acu As String, _
                                      my_CDR As String)

    Dim input_record As String

    Dim my_conta     As Integer

    Dim mytablef     As New ADODB.Recordset

    Dim mysql        As String

    Dim myDato       As String

    Dim myCuenta     As Integer

    On Error GoTo read_save_electronico

    fnum = FreeFile
    Open input_file For Input As #fnum

    Do Until EOF(fnum)
        Line Input #fnum, input_record
 
        my_conta = my_conta + 1
        'aqui llamamos a la base de datos a insertar
        mysql = ""
        mysql = "UPDATE FACTURA " & Chr$(10)
    
        ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
        'If acu = "D" Or acu = "V" Or acu = "C"  Then
        If acu = "D" Or acu = "V" Or acu = "C" Or acu = "E" Or acu = "F" Then
            ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
    
            mysql = mysql & "SET ESTADO_SUNAT='PENDIENTE' , CDR='" & Mid(input_record, 6, 28) & "'," & Chr$(10)
            mysql = mysql & "E_SUNAT= 'N' " & Chr$(10)

        End If

        mysql = mysql & "WHERE serie ='" & Trim("" & my_serie) & "' " & Chr$(10)
        mysql = mysql & "AND numero ='" & Trim("" & my_numero) & "' " & Chr$(10)
        mysql = mysql & "AND tipo = '" & Trim("" & my_tipo) & "' " & Chr$(10)
        'mysql = mysql & "AND E_SUNAT is null " & Chr$(10)
  
        cn.Execute (mysql)

    Loop
    Close #fnum

    Exit Function

read_save_electronico:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

    End Select
    
End Function

Public Function busca_respuesta_electronica(file As String, sFolderPath As String)

    Dim n_file   As String

    Dim sArchivo As String

    Dim my_path  As String

    On Error GoTo busca_respuesta_electronica

    If Mid(sFolderPath, 13, 7) = "FIRMADO" Then
        FileCopy sFolderPath & Left(file, (Len(file) - 10)) & ".txt", "D:\ce_output\FIRMADO\" & "R_" & Left(file, (Len(file) - 10)) & ".txt"
    Else
        FileCopy sFolderPath & Left(file, (Len(file) - 10)) & ".txt", "D:\ce_output\ERROR\" & "R_" & Left(file, (Len(file) - 10)) & ".txt"

    End If

    Exit Function
busca_respuesta_electronica:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

    End Select
 
End Function

Public Function busca_respuesta_sunat(file As String, sFolderPath As String)

    Dim n_file   As String

    Dim sArchivo As String

    Dim my_path  As String

    If Mid(sFolderPath, 13, 9) = "PROCESADO" Then
        FileCopy sFolderPath & Left(file, (Len(file) - 10)) & ".txt", "D:\ce_output\PROCESADO\" & Left(file, (Len(file) - 10))
    Else
        FileCopy sFolderPath & Left(file, (Len(file) - 10)) & ".txt", "D:\ce_output\ERROR\" & Left(file, (Len(file) - 10))

    End If

    Exit Function
 
End Function

Public Function FileExists(sFullPath As String) As Boolean

    Dim oFile As New Scripting.FileSystemObject

    FileExists = oFile.FileExists(sFullPath)

    'MsgBox "FileExists Patricia" & FileExists
End Function

Public Function Busca_comprobante_sunat(my_local As String, _
                                        my_serie As String, _
                                        my_numero As String, _
                                        my_tipo As String, _
                                        salida As Boolean)

    Dim mytable As New ADODB.Recordset

    Dim mysql   As String

    'aqui llamamos a la base de datos a insertar
    mysql = ""
    mysql = "SELECT CDR" & Chr$(10)
    mysql = mysql & "FROM FACTURA" & Chr$(10)
    mysql = mysql & "WHERE local ='" & extra_loquesea(my_local) & "'" & Chr$(10)
    mysql = mysql & "AND serie ='" & Trim("" & my_serie) & "' " & Chr$(10)
    mysql = mysql & "AND numero ='" & Trim("" & my_numero) & "' " & Chr$(10)
    mysql = mysql & "AND tipo = '" & Trim("" & my_tipo) & "' " & Chr$(10)

    mytable.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytable.EOF Then
        Exit Function
    Else
        MY_CELECTRONICO = "" & mytable.Fields("CDR")

        If MY_CELECTRONICO = "" Then
            salida = False
        Else
            salida = True

        End If

    End If

    mytable.Close

    Exit Function

End Function

' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
Public Function Busca_comprobanteRelacionado_sunat(my_local As String, _
                                                   my_serie As String, _
                                                   my_numero As String, _
                                                   my_tipo As String, _
                                                   salida As Boolean)

    Dim mytable As New ADODB.Recordset

    Dim mysql   As String

    mysql = ""
    mysql = "SELECT estado_sunat " & Chr$(10)
    mysql = mysql & "FROM FACTURA" & Chr$(10)
    mysql = mysql & "WHERE local ='" & extra_loquesea(my_local) & "'" & Chr$(10)
    mysql = mysql & "AND serie ='" & Trim("" & my_serie) & "' " & Chr$(10)
    mysql = mysql & "AND numero ='" & Trim("" & my_numero) & "' " & Chr$(10)
    mysql = mysql & "AND tipo = '" & Trim("" & my_tipo) & "' " & Chr$(10)

    mytable.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytable.EOF Then
        MsgBox "VERIFICAR EXISTENCIA DE COMPROBANTE RELACIONADO", vbCritical
        salida = False
        Exit Function
    Else

        If "" & mytable.Fields("estado_sunat") <> "PENDIENTE" Then
            salida = True
        Else
            MsgBox "VERIFICAR ESTADO DE COMPROBANTE RELACIONADO", vbCritical
            salida = False

        End If

    End If

    mytable.Close
    Exit Function

End Function

' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

Function estrae_nota_debito(my_ruc As String, _
                            local1 As String, _
                            bxtipo As String, _
                            bxserie As String, _
                            bxnumero As String, _
                            my_idubigeo As String, _
                            acu As String, _
                            my_struc_datos_empresa() As struc_datos_empresa, _
                            my_struc_ubigeo_Receptor() As struc_ubigeo_Receptor, _
                            my_carga_busca_cliente() As struc_busca_cliente, _
                            my_struc_credito() As struc_credito, _
                            paso As Boolean, _
                            my_struc_ubigeo_Emisor() As struc_ubigeo_Emisor, _
                            file As String, _
                            my_tipo As String)

    ' ' Testing Proyecto Facturacion Electronica  28/02/2018

    Dim mysql           As String

    Dim mytable         As New ADODB.Recordset

    Dim hastaCuanto     As Integer

    Dim nuevoDato       As String

    Dim myDato          As String

    Dim my_precioSinigv As Currency

    Dim myREG           As String

    Dim xnumero         As String

    Dim xserie          As String

    mysql = ""
    mysql = "SELECT d.impuesto as impuestod, d.igv as igvd, f.SERIE," & Chr$(10)
    mysql = mysql & "f.NUMERO, " & Chr$(10) 'serie y número correlativo item1
    mysql = mysql & "CONVERT(VARCHAR(10),f.FECHA, 120) AS FECHA ," & Chr$(10) 'Fecha de emisión item 3
    mysql = mysql & "CONVERT(VARCHAR(10),f.FECHAE, 120) AS FECHAE ," & Chr$(10) 'Fechavencimiento item 4
    mysql = mysql & "f.tipo," & Chr$(10) 'TipoDocumento item 3
    mysql = mysql & "Case f.moneda" & Chr$(10) ' TipoMoneda item 4
    mysql = mysql & "WHEN 'S' THEN 'PEN'" & Chr$(10) ' TipoMoneda item 4
    mysql = mysql & "WHEN 'D' THEN 'USD'" & Chr$(10) ' TipoMoneda item 4
    mysql = mysql & "END AS MONEDA," & Chr$(10) ' TipoMoneda item 4
    '****aqui todo con parametrizacion va al txt
    'Documentos de referencia
    mysql = mysql & "f.adetotal," & Chr$(10) 'serie item 15 'adelanto total
    mysql = mysql & "f.serie1," & Chr$(10) 'serie item 16
    mysql = mysql & "f.numero1," & Chr$(10) 'numero item 16
    mysql = mysql & "f.tipo1," & Chr$(10) 'tipo item 16
    mysql = mysql & "f.tdetra," & Chr$(10) 'tipo item 18
    'aqui los datos de partida y llegada
    mysql = mysql & "f.partida," & Chr$(10)
    mysql = mysql & "f.destino," & Chr$(10)
    'Datos del detalle o Ítem de la Factura
    mysql = mysql & "d.unidad," & Chr$(10)  'UnidadMedidaItem
    mysql = mysql & "d.cantidad," & Chr$(10)  'CantidadItem
    mysql = mysql & "d.producto," & Chr$(10) 'Producto
    mysql = mysql & "d.descripcio as descripcioP," & Chr$(10) 'DescripcionItem producto
    mysql = mysql & "isnull(p.costoinisigv,0) as costoinisigv," & Chr$(10) 'ValorUnitarioSinIgv
    mysql = mysql & "d.precio," & Chr$(10) 'PrecioUnitarioConIgv
    'CodTipoPrecioVtaUnitarioItem 01 precio unitario, 02 valor referencial txt
    mysql = mysql & "f.subtotal," & Chr$(10) 'ImporteIGVItem
    'CodigoAfectacionIGVItem **Afectación al IGV - Catálogo No. 07 txt
    mysql = mysql & "f.tisc," & Chr$(10) 'MontoISCTotal
    mysql = mysql & "d.tisc as dtisc," & Chr$(10) 'MontoISCItem
    'para la posicon rojo 6
    mysql = mysql & "isnull(p.costou,0) as costou," & Chr$(10) 'ValorVentaItem
    'para la posicion 8
    mysql = mysql & "f.impuesto," & Chr$(10) 'ValorVentaItem
    mysql = mysql & "p.igv," & Chr$(10) 'posicion 8
    mysql = mysql & "f.descuento as tdescuento, " & Chr$(10) 'DescuentoItem
    mysql = mysql & "d.descuento as descuento, " & Chr$(10) 'DescuentoItem
    mysql = mysql & "d.cantidad * d.precio as ValorVenta," & Chr$(10) 'ValorVentaItem" y la deduccion de los descuentos
    'Totales de la Boleta de Venta  (Es identico a la Factura) ***
    mysql = mysql & "f.total," & Chr$(10) 'TotalOPGravadas item 30
    mysql = mysql & "isnull(f.fpago,0) as fpago," & Chr$(10)
    mysql = mysql & "d.subtotal," & Chr$(10) 'SumatoriaIGV item 35 **se calcula text
    mysql = mysql & "d.comision as comision," & Chr$(10)
    '**Documentos de referencia/TipoDocumentoRelacionado
    'mysql = mysql & "f.acu," & Chr$(10) 'TipoDocumentoRelacionado posicion 41
    '**Posicion 10/CodTipoPrecioVtaUnitarioItem 01 incluye igv 02 operaciones no honerosas
    mysql = mysql & "f.impuesto as impuesto," & Chr$(10)
    mysql = mysql & "f.subtotal as subtotal," & Chr$(10)
    mysql = mysql & "f.neto as neto," & Chr$(10)
    mysql = mysql & "f.gravado as gravado," & Chr$(10)
 
    '***Información adicional - Percepciones
    mysql = mysql & "f.percepcion," & Chr$(10) 'BaseImponiblePercepcion 41
    mysql = mysql & "d.tpercepcio," & Chr$(10) ' MontoPercepcion 41 *aqui calcular txt
    '***Información Adicional - Factura Guía /Marca y placa de Vehiculo item 52

    ' Testing Proyecto Facturacion Electronica
    mysql = mysql & "f.TIPONCD,f.observa, p.servicio as pservicio" & Chr$(10) ' MontoPercepcion 41 *aqui calcular txt

    ' Testing Proyecto Facturacion Electronica

    mysql = mysql & "from factura f," & Chr$(10)
    mysql = mysql & "detalle d," & Chr$(10)
    mysql = mysql & "producto p" & Chr$(10)
    mysql = mysql & "where f.local='" & "" & local1 & "' " & Chr$(10)
    mysql = mysql & "and f.serie ='" & "" & bxserie & "' " & Chr$(10)
    mysql = mysql & "and f.numero='" & "" & bxnumero & "' " & Chr$(10)
    mysql = mysql & "and f.acu='" & "" & acu & "' " & Chr$(10)
    mysql = mysql & "and f.tipo='" & "" & my_tipo & "' " & Chr$(10)
    mysql = mysql & "and f.SERIE = d.serie" & Chr$(10)
    mysql = mysql & "and f.NUMERO = d.NUMERO" & Chr$(10)
    mysql = mysql & "and p.producto = d.PRODUCTO" & Chr$(10)
    mysql = mysql & "and f.tipo = d.tipo" & Chr$(10)
 
    mytable.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytable.RecordCount > 0 Then
        'para el numero
        hastaCuanto = 8 - Len(mytable.Fields("NUMERO"))
        myDato = mytable.Fields("NUMERO")
        Call E_llenar_zero(hastaCuanto, myDato, xnumero)
        'PARA LA SERIE
        Call E_llenar_zero(3 - Len(mytable.Fields("SERIE")), mytable.Fields("SERIE"), xserie)
 
        If bxtipo = "2" Then
            FileName = "D:\ce_output\CREA\" & my_ruc & "_08" & "_" & xserie & "-" & xnumero & ".INPUT.TXT"
            file = my_ruc & "_08" & "_" & xserie & "-" & xnumero & ".INPUT.TXT"
        ElseIf bxtipo = "1" Then
            FileName = "D:\ce_output\CREA\" & my_ruc & "_08" & "_" & xserie & "-" & xnumero & ".INPUT.TXT"
            file = my_ruc & "_08" & "_" & xserie & "-" & xnumero & ".INPUT.TXT"

        End If
  
        Do
            c = c + 1 'el contador sirve para la lista

            If mytable.EOF Then Exit Do
            'aqui es para abrir el file txt y llenarlo
            Filelibero1 = FreeFile
            Open FileName For Append As #Filelibero1
            conta = conta + 1
      
            conta_record = mytable.RecordCount

            'myREG = ""
            'la cabecera
            If c = 1 Then
                myREG = myREG & "H"
                myREG = myREG & "|"
                'posicion 1
    
                If bxtipo = "2" Then
                    hastaCuanto = 4 - Trim(Len(mytable.Fields("SERIE")))
                    Call E_llenar_zero(3 - Len(mytable.Fields("SERIE")), mytable.Fields("SERIE"), nuevoDato)
                    myREG = myREG & "" & Trim(nuevoDato)
                    myREG = myREG & "-"
                Else
                    hastaCuanto = 4 - Trim(Len(mytable.Fields("SERIE")))
                    Call E_llenar_zero(3 - Len(mytable.Fields("SERIE")), mytable.Fields("SERIE"), nuevoDato)
                    myREG = myREG & "" & Trim(nuevoDato)
                    myREG = myREG & "-"

                End If
     
                'posicion 2
                If Trim(mytable.Fields("NUMERO")) = "vacio" Then
                    myREG = myREG & "|" 'gion separador
                Else
                    hastaCuanto = 8 - Trim$(Len(mytable.Fields("NUMERO")))
                    myDato = mytable.Fields("NUMERO")
                    Call E_llenar_zero(hastaCuanto, myDato, nuevoDato)
                    myREG = myREG & nuevoDato
                    myREG = myREG & "|" 'gion separador

                End If
   
                'posicion 2
                If Trim(mytable.Fields("FECHAE")) = 0 Then
                    Call E_llenar_datos(9, "0", nuevoDato)
                    myREG = myREG & nuevoDato
                    myREG = myREG & "|" 'gion separador
                Else
                    hastaCuanto = 10 - Len(mytable.Fields("FECHAE"))
                    myDato = mytable.Fields("FECHAE")
                    Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'NUMERO 3
                    myREG = myREG & nuevoDato
                    myREG = myREG & "|" 'gion separador

                End If

                '**SerieNumeroModifica
                'posicion 3
   
                ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
                If bxtipo = "2" Then
                    hastaCuanto = 4 - Trim(Len(mytable.Fields("SERIE1")))
                    myDato = mytable.Fields("SERIE1")
                    Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                    myREG = myREG & "" & Trim(nuevoDato)
                    myREG = myREG & "-"
                Else
                    hastaCuanto = 4 - Trim(Len(mytable.Fields("SERIE1")))
                    Call E_llenar_zero(3 - Len(mytable.Fields("SERIE1")), mytable.Fields("SERIE"), nuevoDato)
                    myREG = myREG & "" & Trim(nuevoDato)
                    myREG = myREG & "-"

                End If
      
                If Trim(mytable.Fields("NUMERO1")) = "vacio" Then
                    myREG = myREG & "|" 'gion separador
                Else
                    hastaCuanto = 8 - Trim$(Len(mytable.Fields("NUMERO1"))) '**en la tabla
                    myDato = mytable.Fields("NUMERO1")
                    Call E_llenar_zero(hastaCuanto, myDato, nuevoDato) 'NUMERO 3
                    myREG = myREG & nuevoDato
                    myREG = myREG & "|" 'gion separador

                End If

                ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
 
                'posicion 4
                If bxtipo = "2" Then
                    myREG = myREG & "01"
                    myREG = myREG & "|" 'gion separador
                ElseIf bxtipo = "1" Then
                    myREG = myREG & "03"
                    myREG = myREG & "|" 'gion separador

                End If

                'TipoNotaCredito
                'posicion 5
  
                If mytable.Fields("TIPONCD") = Null Then 'posicion 6
                    myREG = myREG & "|"
                Else
                    myREG = myREG & mytable.Fields("TIPONCD")
                    myREG = myREG & "|" 'gion separador

                End If

                'MotivoSustento
                'posicion 6
                If mytable.Fields("observa") = Null Then
                    myREG = myREG & "|"
                Else
                    Call llenar_datos(250 - Trim(Len(mytable.Fields("observa"))), mytable.Fields("observa"), nuevoDato) 'tipoclie 4
                    myREG = myREG & Trim(nuevoDato)
                    myREG = myREG & "|" 'gion separador

                End If

                'posicion 7
                If Trim(mytable.Fields("MONEDA")) = Null Then
                    myREG = myREG & "|"
                Else
                    Call llenar_datos(3 - Len(mytable.Fields("MONEDA")), mytable.Fields("MONEDA"), nuevoDato)  'MONEDA 9
                    myREG = myREG & nuevoDato
                    myREG = myREG & "|" 'gion separador

                End If

                'Datos del Emisor **NumeroDocIdentidadReceptor 'Ruc
                'posicion 8
      
                If my_carga_busca_cliente(0).RUC = "" Then
                    myREG = myREG & "000000|"
                Else
                    Call E_llenar_datos(11 - Trim(Len(my_carga_busca_cliente(0).RUC)), Trim(my_carga_busca_cliente(0).RUC), nuevoDato) 'posicion 5
                    myREG = myREG & Trim(nuevoDato) 'item 12
                    myREG = myREG & "|" 'gion separador

                End If

                If Len(my_carga_busca_cliente(0).RUC) = 11 Then
                    myREG = myREG & "6" 'REG. UNICO DE CONTRIBUYENTES
                    myREG = myREG & "|"
                ElseIf Len(my_carga_busca_cliente(0).RUC) = 8 Then
                    myREG = myREG & "1" 'DOC. NACIONAL DE IDENTIDAD
                    myREG = myREG & "|"
                ElseIf Len(my_carga_busca_cliente(0).RUC) = 9 Then
                    myREG = myREG & "4" 'CARNET DE EXTRANJERIA
                    myREG = myREG & "|"
                Else
                    myREG = myREG & "0" 'OTROS
                    myREG = myREG & "|"

                End If

                '**Datos del cliente o receptor/RazonSocialReceptor
                'posicion 6
                If my_carga_busca_cliente(0).nombre = "" Then
                    myREG = myREG & "PUBLICO General|"
                Else
                    myREG = myREG & Trim(my_carga_busca_cliente(0).nombre)
                    myREG = myREG & "|"

                End If

                myREG = myREG & Trim(my_struc_ubigeo_Receptor(0).idubigeo)
                myREG = myREG & "|" 'gion separador

                '**Datos del cliente o receptor/DireccionReceptor
                'posicion 7
                myREG = myREG & Trim(my_carga_busca_cliente(0).direccion)
                myREG = myREG & "|" 'gion separador
   
                myREG = myREG & ""
                myREG = myREG & "|" 'gion separador
 
                myREG = myREG & ""
                myREG = myREG & "|" 'gion separador
   
                myREG = myREG & Trim(my_carga_busca_cliente(0).dpto)
                myREG = myREG & "|" 'gion separador
   
                myREG = myREG & ""
                myREG = myREG & "|" 'gion separador
     
                '**Dirección del cliente/CodigoPaisReceptor
                'posicion 17
                myREG = myREG & "PE" ' muestra de b&H
                myREG = myREG & "|" 'gion separador

                '**Datos del cliente o receptor/CorreoReceptor
                'posicion 18
                Call E_llenar_datos(100 - Len(my_carga_busca_cliente(0).correo), my_carga_busca_cliente(0).correo, nuevoDato)
                myREG = myREG & Trim(nuevoDato) 'item 15
                myREG = myREG & "|" 'gion separador
                ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

                ' Testing Proyecto Facturacion Electronica
                '**TotalOPGravadas
                'posicion 19
                Call E_llenar_datos(15 - Len(mytable.Fields("subtotal")), mytable.Fields("subtotal"), nuevoDato)
                nuevoDato = nuevoDato - mytable.Fields("gravado")
                myREG = myREG & Format(Trim(nuevoDato), "0.00")
                myREG = myREG & "|" 'gion separador 'es junto con bajo
                ' Testing Proyecto Facturacion Electronica

                '**TotalOPNoGravadas
                'posicion 20
                If my_struc_datos_empresa(0).Toperacion = "I" Then
                    Call E_llenar_datos(15 - Len(mytable.Fields("gravado")), mytable.Fields("gravado"), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|" 'gion separador
                Else
                    myREG = myREG & "|" 'gion separador

                End If

                '**TotalOPExonerado
                'posicion 21
                If my_struc_datos_empresa(0).Toperacion = "E" Then
                    Call E_llenar_datos(15 - Len(mytable.Fields("gravado")), mytable.Fields("gravado"), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|" 'gion separador 'es junto con bajo
                Else
                    myREG = myREG & "|" 'gion separador

                End If
 
                If "" & mytable.Fields("tipo") = "7" Then
                    myREG = myREG & "|" 'gion separador
                Else
                    Call E_llenar_datos(15 - Trim(Len(mytable.Fields("impuesto"))), mytable.Fields("impuesto"), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|" 'gion separador

                End If
  
                '**SumatoriaISC
                'posicion 24
                If "" & mytable.Fields("tisc") = "0" Then
                    myREG = myREG & "|"
                Else
                    Call E_llenar_datos(15 - Trim(Len("" & mytable.Fields("tisc"))), Trim("" & mytable.Fields("tisc")), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|" 'gion separador

                End If
  
                '**SumatoriaOtrosTributos
                'posicion 25
                If Trim(mytable.Fields("tdetra")) = "0" Then
                    myREG = myREG & "|"
                Else
                    Call E_llenar_datos(15 - Trim(Len(mytable.Fields("tdetra"))), Trim(mytable.Fields("tdetra")), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|" 'gion separador

                End If
    
                '**SumatoriaOtrosCargos
                'posicion 26
                If Trim("" & mytable.Fields("comision")) = "0" Then
                    myREG = myREG & "|"
                Else
                    Call E_llenar_datos(15 - Trim(Len("" & mytable.Fields("comision"))), "" & mytable.Fields("comision"), nuevoDato)
                    myREG = myREG & Format(Trim(nuevoDato), "0.00")
                    myREG = myREG & "|" 'gion separador

                End If

                '**ImporteTotal
                'posicion 27
                Call E_llenar_datos(15 - Trim(Len(mytable.Fields("total"))), mytable.Fields("total"), nuevoDato)
                myREG = myREG & Format(Trim(Round(nuevoDato, 2)), "0.00")
                myREG = myREG & "|"
                myREG = myREG & "|"

                Print #Filelibero1, myREG
                'para el detalle del producto
 
                myREG = myREG & Chr(13)

            End If

            myREG = ""
            myREG = myREG & "I"
            myREG = myREG & "|"
        
            '**Datos del detalle
            'posicion 1 OrdenItem
            myDato = c

            If Len(myDato) = 1 Then
                myREG = myREG & "0" & myDato
                myREG = myREG & "|"
            Else
                myREG = myREG & myDato
                myREG = myREG & "|"

            End If

            'posicion 2 UnidadMedidaItem
            If mytable.Fields("unidad") = Null Then
                myREG = myREG & "|"
            Else

                If mytable.Fields("pservicio") = "S" Then
                    myREG = myREG & "ZZ|"
                Else
                    myREG = myREG & "NIU|"

                End If

            End If
            
            'posicion 3 CantidadItem
            If mytable.Fields("cantidad") = Null Then
                myREG = myREG & "|"
            Else
                Call E_llenar_datos(23 - Trim(Len(mytable.Fields("cantidad"))), mytable.Fields("cantidad"), nuevoDato)
                myREG = myREG & Trim(nuevoDato)
                myREG = myREG & "|"

            End If
            
            'posicion 4 CodigoProductoItem
            If mytable.Fields("producto") = Null Then
                myREG = myREG & "|"
            Else
                Call llenar_datos(30 - Trim(Len(mytable.Fields("producto"))), mytable.Fields("producto"), nuevoDato)
                myREG = myREG & Trim(nuevoDato)
                myREG = myREG & "|"

            End If

            'posicion 5 DescripcionItem
            If mytable.Fields("descripcioP") = Null Then
                myREG = myREG & "|"
            Else
                Call E_llenar_datos(250 - Trim(Len(mytable.Fields("descripcioP"))), mytable.Fields("descripcioP"), nuevoDato)
                myREG = myREG & Trim(nuevoDato)
                myREG = myREG & "|"

            End If
    
            'posicion 6 ValorUnitarioSinIgv
            my_precioSinigv = 0
            my_precioSinigv = (mytable.Fields("precio") / (1 + mytable.Fields("igv") / 100))
            my_dato = Round(my_precioSinigv, 2)
            myREG = myREG & Format(Trim(my_dato), "0.00")
            myREG = myREG & "|"
  
            'posicion 7 PrecioUnitarioConIgv
            Call E_llenar_datos(15 - Len(mytable.Fields("precio")), mytable.Fields("precio"), nuevoDato)  'item 28
            myREG = myREG & Format(Trim(nuevoDato), "0.00")
            myREG = myREG & "|" 'gion separador 'es junto con bajo
  
            'posicion 8 CodTipoPrecioVtaUnitarioItem
            myREG = myREG & "01"
            myREG = myREG & "|"
            
            'posicion 9 ImporteIGVItem
            '''' 17/07/2018 Factura de Exportación
            Call E_llenar_datos(15 - Trim(Len(mytable.Fields("impuestod"))), mytable.Fields("impuestod"), nuevoDato) 'item 27
            myREG = myREG & Format(Trim(nuevoDato), "0.00")
            myREG = myREG & "|"
            '''' 17/07/2018 Factura de Exportación
            
            'posicion 10 CodigoAfectacionIGVItem
            '''' 17/07/2018 Factura de Exportación
            If mytable.Fields("igvd") = "0" Then
                If my_tipoigv = "40" Then
                    myREG = myREG & "40|"
                Else
                    myREG = myREG & "20|"

                End If

            Else
                myREG = myREG & "10|"

            End If

            'posicion 11 MontoISCItem
            If Trim("" & mytable.Fields("dtisc")) = "" Or Trim("" & mytable.Fields("dtisc")) = Null Or Trim("" & mytable.Fields("dtisc")) = "0" Then
                myREG = myREG & "|"
            Else
                Call E_llenar_datos(15 - Trim(Len(mytable.Fields("dtisc"))), Trim(mytable.Fields("dtisc")), nuevoDato)
                myREG = myREG & Format(Trim(nuevoDato), "0.00")
                myREG = myREG & "|"

            End If

            'posicion 12 TipoSistemaISCItem
            myREG = myREG & "03|"

            'posicion 13 ValorVentaItem
            my_precioSinigv = mytable.Fields("precio") / (1 + mytable.Fields("igv") / 100)
            'my_dato = Round(my_precioSinigv, 2) * mytable.Fields("cantidad")
            my_dato = Round(my_precioSinigv, 2) * mytable.Fields("cantidad") - mytable.Fields("descuento")
            myREG = myREG & Format(Trim(my_dato), "0.00")
  
            Print #Filelibero1, myREG
   
            ' Campos Adicionales FE 19/05/2018
            If c = conta_record Then
                myREG = myREG & Chr(13)
                myREG = ""
                myREG = myREG & "A"
                myREG = myREG & "|1||OBSERVACIÓN: "
                myREG = myREG & mytable.Fields("observa")
                Print #Filelibero1, myREG

            End If

            '  myREG = myREG & Chr(13)
            '   myREG = ""
            '   myREG = myREG & "A"
            '   myREG = myREG & "|1|100|Orden:ver-4"
            '   Print #Filelibero1, myREG
            '
            '   myREG = myREG & "A"
            '    myREG = myREG & Chr(13)
            '   myREG = ""
            '   myREG = myREG & "|1|100|xxxx"
            '   Print #Filelibero1, myREG
   
            ' Campos Adicionales FE 19/05/2018

            Close #Filelibero1
            mytable.MoveNext
        Loop
        Close #Filelibero1

    End If

    mytable.Close

    'aqui copia el archivo txt

    origen = FileName

    destino = "D:\ce_Input\" & file

    ' Testing Proyecto Facturacion Electronica 28/02/2018
    'FileCopy origen, destino
    Dim fso As New Scripting.FileSystemObject

    fso.MoveFile origen, destino
    ' Testing Proyecto Facturacion Electronica  28/02/2018

End Function

'inicio 22/12/2107 pll
Function estrae_baja(my_ruc As String, _
                     my_local As String, _
                     my_tipo As String, _
                     my_serie As String, _
                     my_numero As String, _
                     my_fecha As String, _
                     file As String, _
                     salida As Boolean)

    Dim mysql       As String

    Dim mytable     As New ADODB.Recordset

    Dim hastaCuanto As Integer

    Dim nuevoDato   As String

    Dim myDato      As String

    On Error GoTo estrae_baja

    mysql = ""
    'Datos de la Boleta de Venta (Es identico a la factura)
    mysql = "SELECT distinct f.SERIE," & Chr$(10)
    mysql = mysql & "f.NUMERO, " & Chr$(10) 'serie y número correlativo item1
    mysql = mysql & "CONVERT(VARCHAR(10),f.FECHAE, 120) AS FECHAE ," & Chr$(10) 'Fechavencimiento item 4
    mysql = mysql & "f.tipo" & Chr$(10) 'TipoDocumento item 3
    mysql = mysql & "from factura f," & Chr$(10)
    mysql = mysql & "detalle d," & Chr$(10)
    mysql = mysql & "fpagov p" & Chr$(10)
    mysql = mysql & "where f.local ='" & "" & my_local & "' " & Chr$(10)
    mysql = mysql & "and f.tipo='" & "" & my_tipo & "' " & Chr$(10)
    mysql = mysql & "and f.serie ='" & "" & my_serie & "' " & Chr$(10)
    mysql = mysql & "and f.numero='" & "" & my_numero & "' " & Chr$(10)
    mysql = mysql & "and f.tipo= d.tipo" & Chr$(10)
    mysql = mysql & "and f.serie = d.serie" & Chr$(10)
    mysql = mysql & "and f.numero= d.numero" & Chr$(10)
    mysql = mysql & "and f.tipo= p.tipo" & Chr$(10)
    mysql = mysql & "and f.serie = p.serie" & Chr$(10)
    mysql = mysql & "and f.numero= p.numero" & Chr$(10)
    mysql = mysql & "and f.estado='1'" & Chr$(10)
    mysql = mysql & "and d.estado='1'" & Chr$(10)
    mysql = mysql & "and p.estado='1'" & Chr$(10)
    
    mytable.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytable.RecordCount > 0 Then
        salida = True
        'para el numero
        hastaCuanto = 8 - Len(mytable.Fields("NUMERO")) '**en la tabla
        myDato = mytable.Fields("NUMERO")
        Call E_llenar_zero(hastaCuanto, myDato, my_numero)
        'PARA LA SERIE
 
        Call E_llenar_zero(4 - Len(mytable.Fields("SERIE")), mytable.Fields("SERIE"), my_serie)

        ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
        ' If my_tipo = "2" Then
        '   FileName = "D:\ce_output\CREA\" & my_ruc & "_RA" & "_" & my_serie & "-" & my_numero & ".INPUT.TXT"
        '   file = my_ruc & "_RA" & "_" & my_serie & "-" & my_numero & ".INPUT.TXT"
        '
        '  ElseIf my_tipo = "1" Then
        '    FileName = "D:\ce_output\CREA\" & my_ruc & "_RA" & "_" & my_serie & "-" & my_numero & ".INPUT.TXT"
        '    file = my_ruc & "_RA" & "_" & my_serie & "-" & my_numero & ".INPUT.TXT"
        '  End If
        '
        ' If my_tipo1 = "1" Or my_tipo1 = "2" Then ' nc
        '     FileName = "D:\ce_output\CREA\" & my_ruc & "_RA" & "_" & my_serie & "-" & my_numero & ".INPUT.TXT"
        '     file = my_ruc & "_RA" & "_" & my_serie & "-" & my_numero & ".INPUT.TXT"
        ' End If
 
        FileName = "D:\ce_output\CREA\" & my_ruc & "_RA" & "_" & my_serie & "-" & my_numero & ".INPUT.TXT"
        file = my_ruc & "_RA" & "_" & my_serie & "-" & my_numero & ".INPUT.TXT"
        ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
 
        Do
            c = c + 1 'el contador sirve para la lista

            If mytable.EOF Then Exit Do
            'aqui es para abrir el file txt y llenarlo
            Filelibero1 = FreeFile
            Open FileName For Append As #Filelibero1
            conta = conta + 1
      
            conta_record = mytable.RecordCount

            'myREG = ""
            'la cabecera
            If c = 1 Then
                myREG = myREG & "H"
                myREG = myREG & "|"
      
                '**fecha emision
                If Trim(mytable.Fields("FECHAE")) = 0 Then
                    Call E_llenar_datos(9, "0", nuevoDato)
                    myREG = myREG & nuevoDato
                    myREG = myREG & "|" 'gion separador
                Else
                    hastaCuanto = 10 - Len(mytable.Fields("FECHAE"))
                    myDato = mytable.Fields("FECHAE")
                    Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'NUMERO 3
                    myREG = myREG & nuevoDato
                    myREG = myREG & "|" 'gion separador

                End If
   
                Print #Filelibero1, myREG
                'para el detalle del producto
 
                myREG = myREG & Chr(13)

            End If

            myREG = ""
            myREG = myREG & "I"
            myREG = myREG & "|"

            '****inicio
            '***Datos del detalle o Ítem (Es identico a la factura)/Número de orden del Ítem
            'myDato = c 'Número de orden del Ítem 'item 18
            'tipo documento posicion1
            If my_tipo = "2" Then 'item 19 posicion 4 rojo
                myREG = myREG & "01|"
            ElseIf my_tipo = "1" Then
                myREG = myREG & "03|"
            ElseIf my_tipo1 = "2" Then

                If my_acu = "E" Then 'NC
                    myREG = myREG & "07|"
                ElseIf my_acu = "F" Then 'NC
                    myREG = myREG & "08|"

                End If

            ElseIf my_tipo1 = "1" Then

                If my_acu = "E" Then 'NC
                    myREG = myREG & "07|"
                ElseIf my_acu = "F" Then 'NC
                    myREG = myREG & "08|"

                End If

            End If
 
            '**serie documento posicion 2
            '**Datos del detalle o Ítem (Es identico a la factura)/UnidadMedidaItem
 
            ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

            If mytable.Fields("serie") = Null Then 'item 19 posicion 4 rojo
                myREG = myREG & "|"
            Else 'item 18 item 1 color rojo
                Call E_llenar_datos(4 - Trim(Len(mytable.Fields("serie"))), mytable.Fields("serie"), nuevoDato)
                myREG = myREG & "" & Trim(nuevoDato)
                myREG = myREG & "|" 'gion separador

            End If

            ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
 
            '**numero /posicion 3
            If Trim(mytable.Fields("NUMERO")) = Null Then
                myREG = myREG & "|" 'gion separador
            Else
                hastaCuanto = 8 - Trim$(Len(mytable.Fields("NUMERO"))) '**en la tabla
                myDato = mytable.Fields("NUMERO")
                Call E_llenar_zero(hastaCuanto, myDato, nuevoDato) 'NUMERO 3
                myREG = myREG & nuevoDato
                myREG = myREG & "|" 'gion separador

            End If

            '///
            '*motivo de baja posicion 4 el sistema no lo tiene
            myREG = myREG & "BAJA|"
   
            Print #Filelibero1, myREG
            Close #Filelibero1
            mytable.MoveNext
        Loop
        Close #Filelibero1

    End If

    mytable.Close

    'aqui copia el archivo txt

    origen = FileName

    destino = "D:\ce_Input\" & file

    ' Testing Proyecto Facturacion Electronica 28/02/2018
    'FileCopy origen, destino
    Dim fso As New Scripting.FileSystemObject

    fso.MoveFile origen, destino
    ' Testing Proyecto Facturacion Electronica  28/02/2018

    Exit Function

estrae_baja:

    Select Case Err.Number

        Case 71

            Resume Next

        Case 424

            Resume Next

        Case Else

            Resume Next

    End Select
 
End Function

'fin 22/12/2017 pll
Public Function busca_electronico(my_local As String, _
                                  my_serie As String, _
                                  my_numero As String, _
                                  my_tipo As String, _
                                  my_CDR As String)

    Dim mytable As New ADODB.Recordset

    Dim mysql   As String

    'aqui llamamos a la base de datos a insertar
    mysql = ""
    mysql = "SELECT CDR" & Chr$(10)
    mysql = mysql & "FROM FACTURA " & Chr$(10)
    mysql = mysql & "WHERE LOCAL ='" & Trim("" & my_local) & "' " & Chr$(10)
    mysql = mysql & "AND serie ='" & Trim("" & my_serie) & "' " & Chr$(10)
    mysql = mysql & "AND numero ='" & Trim("" & my_numero) & "' " & Chr$(10)
    mysql = mysql & "AND tipo = '" & Trim("" & my_tipo) & "' " & Chr$(10)
 
    mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
 
    If mytable.EOF Then
        salida = False
        Exit Function
    Else
        salida = True
        my_CDR = mytable.Fields("CDR")

    End If
  
    Exit Function
    
End Function

Public Function Actualiza_cdr(input_file As String, _
                              my_local As String, _
                              my_serie As String, _
                              my_numero As String, _
                              my_tipo As String, _
                              acu As String, _
                              my_CDR As String)

    Dim input_record As String

    Dim my_conta     As Integer

    Dim mytablef     As New ADODB.Recordset

    Dim mysql        As String

    Dim myDato       As String

    Dim myCuenta     As Integer

    'On Error GoTo Actualiza_cdr

    fnum = FreeFile
    Open input_file For Input As #fnum

    Do Until EOF(fnum)
    
        Line Input #fnum, input_record
 
        my_conta = my_conta + 1
        'aqui llamamos a la base de datos a insertar
        mysql = "UPDATE COMPROBANTEE " & Chr$(10)

        If acu = "E" Then 'Nota de credito venta
            mysql = mysql & "SET CDR_NCV='" & Mid(input_record, 6, 28) & "'" & Chr$(10)

        End If

        If acu = "F" Then 'Nota de credito venta
            mysql = mysql & "SET CDR_NDV='" & Mid(input_record, 6, 28) & "'" & Chr$(10)

        End If

        mysql = mysql & "WHERE serie ='" & Trim("" & my_serie) & "' " & Chr$(10)
        mysql = mysql & "AND numero ='" & Trim("" & my_numero) & "' " & Chr$(10)
        mysql = mysql & "AND tipo = '" & Trim("" & my_tipo) & "' " & Chr$(10)

        cn.Execute (mysql)
    
        'para la factura
        mysql = ""
        mysql = "UPDATE FACTURA" & Chr$(10)

        If acu = "E" Then 'Nota de credito venta
            mysql = mysql & "SET CDR_NCV='" & Mid(input_record, 6, 28) & "'" & Chr$(10)

        End If

        If acu = "F" Then 'Nota de credito venta
            mysql = mysql & "SET CDR_NDV='" & Mid(input_record, 6, 28) & "'" & Chr$(10)

        End If

        mysql = mysql & "WHERE serie ='" & Trim("" & my_serie) & "' " & Chr$(10)
        mysql = mysql & "AND numero ='" & Trim("" & my_numero) & "' " & Chr$(10)
        mysql = mysql & "AND tipo = '" & Trim("" & my_tipo) & "' " & Chr$(10)
    
        cn.Execute (mysql)
    
        If acu = "E" Then 'Nota de credito venta
            mysql = ""
            mysql = "UPDATE FACTURA" & Chr$(10)
            mysql = mysql & "SET CDR_NCV='" & Mid(input_record, 6, 28) & "'" & Chr$(10)
            mysql = mysql & "WHERE serie1='" & Trim("" & my_serie1) & "' " & Chr$(10)
            mysql = mysql & "AND numero1 ='" & Trim("" & my_numero1) & "' " & Chr$(10)
            mysql = mysql & "AND tipo = 'NCV' " & Chr$(10)

        End If
   
        If acu = "F" Then 'Nota de credito venta
            mysql = ""
            mysql = "UPDATE FACTURA" & Chr$(10)
            mysql = mysql & "SET CDR_NDV='" & Mid(input_record, 6, 28) & "'" & Chr$(10)
            mysql = mysql & "WHERE serie1='" & Trim("" & my_serie1) & "' " & Chr$(10)
            mysql = mysql & "AND numero1 ='" & Trim("" & my_numero1) & "' " & Chr$(10)
            mysql = mysql & "AND tipo = 'NDV' " & Chr$(10)

        End If

        cn.Execute (mysql)
  
    Loop
    Close #fnum
    Exit Function
    
End Function

Public Function bkp_en_crear()

    Dim sNextFile As String

    On Error GoTo bkp_en_crear

    my_fecha = Format(Now, "dd_MM_yyyy")

    If fso.FileExists("D:\ce_output\CREA\bck_") & my_fecha Then
        MkDir ("D:\ce_output\CREA\bck_") & my_fecha
    Else
        MkDir ("D:\ce_output\CREA\bck_") & my_fecha

    End If

    PathDestino = ("D:\ce_output\CREA\bck_") & my_fecha & "\"

    PathOrigen = "D:\ce_output\CREA\"

    sNextFile = Dir$("D:\ce_output\CREA\*.txt")

    While sNextFile <> ""

        ArchivoOrigen = PathOrigen & sNextFile
        ArchivoDestino = PathDestino & sNextFile
        FileCopy ArchivoOrigen, ArchivoDestino
        sNextFile = Dir$
        'Para eliminar
        Kill ArchivoOrigen
    Wend

    MsgBox "termino bck_crea"
    Exit Function

bkp_en_crear:

    Select Case Err.Number

        Case 75

            Resume Next

        Case 424

            Resume Next

        Case Else

            Resume Next

    End Select

End Function

Public Function bck_en_d_firmado_envia_POR_ENVIAR()

    Dim sNextFile As String

    On Error GoTo bck_en_d_firmado_envia_POR_ENVIAR

    'aqui efectuamos la copia de firmado a enviar solo .xml

    PathOrigen = "D:\ce_Input\FIRMADO\"

    PathDestino = "D:\ce_Input\POR_ENVIAR\"

    sNextFile = Dir$("D:\ce_Input\FIRMADO\*.xml")

    While sNextFile <> ""

        ArchivoOrigen = PathOrigen & sNextFile
        ArchivoDestino = PathDestino & sNextFile
        FileCopy ArchivoOrigen, ArchivoDestino
        sNextFile = Dir$
    Wend

    'aqui efectuamos copia de seguridad
    PathOrigen = ""
    PathDestino = ""
    sNextFile = ""

    my_fecha = Format(Now, "dd_MM_yyyy")

    MkDir ("D:\ce_output\FIRMADO\bck_") & my_fecha

    If FileExists("D:\ce_output\FIRMADO\bck_") & my_fecha Then
        MkDir ("D:\ce_output\FIRMADO\bck_") & my_fecha
    Else
        MkDir ("D:\ce_output\FIRMADO\bck_") & my_fecha

    End If

    PathOrigen = "D:\ce_Input\FIRMADO\"
                                                                                                                                       
    PathDestino = "D:\ce_output\FIRMADO\bck_" & my_fecha & "\"

    sNextFile = Dir$("D:\ce_Input\FIRMADO\*.*")

    While sNextFile <> ""

        ArchivoOrigen = PathOrigen & sNextFile
        ArchivoDestino = PathDestino & sNextFile
        FileCopy ArchivoOrigen, ArchivoDestino
        sNextFile = Dir$
        'Para eliminar
        Kill ArchivoOrigen 'por mientras no
    Wend

    Exit Function

bck_en_d_firmado_envia_POR_ENVIAR:

    Select Case Err.Number

        Case 75

            Resume Next

        Case 424

            Resume Next

        Case Else

            Resume Next

    End Select

End Function

Public Function bck_en_POR_ENVIAR()

    Dim sNextFile As String

    On Error GoTo bck_en_POR_ENVIAR

    my_fecha = Format(Now, "dd_MM_yyyy")

    If FileExists("D:\ce_Input\POR_ENVIAR\bck_") & my_fecha Then
        MkDir ("D:\ce_Input\POR_ENVIAR\bck_") & my_fecha
    Else
        MkDir ("D:\ce_Input\POR_ENVIAR\bck_") & my_fecha

    End If
 
    my_bck = ("D:\ce_Input\POR_ENVIAR\bck_") & my_fecha & "_" & c

    PathOrigen = "D:\ce_Input\POR_ENVIAR\"
    PathDestino = my_bck & "\"

    NEW_PathDestino = "D:\ce_Input\POR_ENVIAR\"
    sNextFile = Dir$("D:\ce_Input\POR_ENVIAR\*.*")

    While sNextFile <> ""

        ArchivoOrigen = PathOrigen & sNextFile
        ArchivoDestino = PathDestino & sNextFile
        ArchivoDestino2 = NEW_PathDestino & sNextFile
        FileCopy ArchivoOrigen, ArchivoDestino
        FileCopy ArchivoOrigen, ArchivoDestino2
        sNextFile = Dir$
        'Para eliminar
        Kill ArchivoOrigen
    Wend

bck_en_POR_ENVIAR:

    Select Case Err.Number

        Case 75

            Resume Next

        Case Else

            Resume Next

    End Select

End Function

Public Function bkp_en_out_FIRMADO()

    Dim sNextFile As String

    On Error GoTo bkp_en_out_FIRMADO

    my_fecha = Format(Now, "dd_MM_yyyy")

    If fso.FileExists("D:\ce_output\FIRMADO\bck_") & my_fecha Then
        MkDir ("D:\ce_output\FIRMADO\bck_") & my_fecha
    Else
        MkDir ("D:\ce_output\FIRMADO\bck_") & my_fecha

    End If

    my_bck = ("D:\ce_output\FIRMADO\bck_") & my_fecha

    PathOrigen = "D:\ce_output\FIRMADO\"
    PathDestino = my_bck & "\"

    sNextFile = Dir$("D:\ce_output\FIRMADO\*.*")

    While sNextFile <> ""

        ArchivoOrigen = PathOrigen & sNextFile
        ArchivoDestino = PathDestino & sNextFile
        FileCopy ArchivoOrigen, ArchivoDestino
        sNextFile = Dir$
        'Para eliminar
        Kill ArchivoOrigen
    Wend

    Exit Function

bkp_en_out_FIRMADO:

    Select Case Err.Number

        Case 75

            Resume Next

        Case 424

            Resume Next

        Case Else

            Resume Next

    End Select

End Function

Public Function bkp_en_out_PROCESADO()

    Dim sNextFile As String

    On Error GoTo bkp_en_out_PROCESADO

    my_fecha = Format(Now, "dd_MM_yyyy")

    If FileExists("D:\ce_output\PROCESADO\bck_") & my_fecha Then
        MkDir ("D:\ce_output\PROCESADO\bck_") & my_fecha
    Else
        MkDir ("D:\ce_output\PROCESADO\bck_") & my_fecha

    End If

    PathOrigen = "D:\ce_input\PROCESADO\"

    PathDestino = "D:\ce_output\PROCESADO\bck_" & my_fecha & "\"

    sNextFile = Dir$("D:\ce_output\PROCESADO\*.*")

    While sNextFile <> ""

        ArchivoOrigen = PathOrigen & sNextFile
        ArchivoDestino = PathDestino & sNextFile
        FileCopy ArchivoOrigen, ArchivoDestino
        sNextFile = Dir$
        'Para eliminar
        Kill ArchivoOrigen
    Wend
    MsgBox "termino bck_procesado"
    Exit Function
bkp_en_out_PROCESADO:

    Select Case Err.Number

        Case 75

            Resume Next

        Case 424

            Resume Next

        Case Else

            Resume Next

    End Select

End Function

'Public Function lee_conf_BYH(input_file As String, new_ruc As String, my_ondemand)
Public Function lee_conf_BYH(input_file As String, my_ondemand)

    Dim input_record  As String

    Dim my_ruc        As String

    Dim my_EnvioSunat As String

    Dim f             As Integer

    Dim Contenido, ContenidoFinal            As String

    archivo = FreeFile
    sRuc = "numeroDocumentoEmisor:"
    sEnvioSunat = "on-demand:"

    f = FreeFile
      
    'aqui abre el file de configuracion para el ruc
    Open input_file For Input As f
      
    Do While Not EOF(f)
        Line Input #f, linea

        '      'aqui busca el ruc
        If InStr(linea, sRuc) <> 0 Then
            my_ruc = Trim(Right(linea, 11)) '

        End If

        'aqui busca envio sunat
        If InStr(linea, sEnvioSunat) <> 0 Then
            'my_EnvioSunat = Trim(Right(linea, 5))
            my_EnvioSunat = linea

        End If

    Loop
    '
    '    'Cierra el archivo
    Close #f
    '
    Open input_file For Input As f
    '
    '   'carga el contenido del archivo en la variable
    Contenido = input$(LOF(f), #f)
    '
    ''Cierra el archivo
    Close #f
    '
    ' 'aqui reemplaza ruc
    '  Contenido = Replace(Contenido, Trim(my_ruc), Trim(new_ruc))

    'aqui reemplaza modalidad sunat
    If my_ondemand = "A" Then
        ContenidoFinal = Replace(Contenido, Trim(my_EnvioSunat), "on-demand: false")
    Else
        ContenidoFinal = Replace(Contenido, Trim(my_EnvioSunat), "on-demand: true")

    End If

    'squi escribe todo ruc y ondemand
    f = FreeFile
    'Abre un nuevo archivo
    Open input_file For Output As f
    'Graba los nuevos datos
    'Print #F, Contenido
    Print #f, ContenidoFinal
    'cierra el archivo
    Close #f
    
    Exit Function

End Function

'Testing Proyecto Facturacion Electronica 05/04/2018
' Verifique RUC
Public Function lee_conf_RUC(input_file As String, my_ondemand)

    Dim input_record  As String

    Dim my_ruc        As String

    Dim my_EnvioSunat As String

    Dim f             As Integer

    Dim Contenido, ContenidoFinal            As String

    archivo = FreeFile
    sRuc = "numeroDocumentoEmisor:"
    sEnvioSunat = "on-demand:"

    f = FreeFile
      
    Open input_file For Input As f
      
    Do While Not EOF(f)
        Line Input #f, linea

        If InStr(linea, sRuc) <> 0 Then
            my_ruc = Trim(Right(linea, 11)) '

        End If

    Loop
    
    Close #f

    Open input_file For Input As f

    Contenido = input$(LOF(f), #f)
    Close #f
 
    'If treevuti.Visible = True Then
    treevuti.RucYml = my_ruc
    ' End If
 
    'If menup.Visible = True Then
    '    menup.RucYml = my_ruc
    'End If

    If my_ondemand = "A" Then
        ContenidoFinal = Replace(Contenido, Trim(my_EnvioSunat), "on-demand: false")
    Else
        ContenidoFinal = Replace(Contenido, Trim(my_EnvioSunat), "on-demand: false")

    End If

    f = FreeFile
    Open input_file For Output As f
    Print #f, ContenidoFinal
    Close #f
    Exit Function

End Function

'Testing Proyecto Facturacion Electronica 05/04/2018

Public Function Graba_RSunat(input_file As String, _
                             my_serie As String, _
                             my_numero As String, _
                             my_tipo As String, _
                             my_caja)

    Dim input_record As String

    Dim my_conta     As Integer

    Dim mytablef     As New ADODB.Recordset

    Dim mysql        As String

    Dim myDato       As String

    Dim myCuenta     As Integer

    fnum = FreeFile
    Open input_file For Input As #fnum

    Do Until EOF(fnum)
    
        Line Input #fnum, input_record
 
        my_conta = my_conta + 1
        'aqui llamamos a la base de datos a insertar
        mysql = ""
        mysql = "UPDATE FACTURA " & Chr$(10)

        If Mid(input_record, 3, 2) = "OK" Then
            mysql = mysql & "SET E_SUNAT='S'" & Chr$(10)
        Else
            mysql = mysql & "SET E_SUNAT='" & Mid(input_record, 3, 2) & "'" & Chr$(10)

        End If

        If my_caja = "00" Then
            mysql = mysql & "WHERE serie ='" & Trim("" & my_serie) & "' " & Chr$(10)
            mysql = mysql & "AND numero ='" & Trim("" & my_numero) & "' " & Chr$(10)
            mysql = mysql & "AND tipo = '" & Trim("" & my_tipo) & "' " & Chr$(10)
        Else
            mysql = mysql & "WHERE LOCAL =" & Trim("" & my_caja) & " " & Chr$(10)
            mysql = mysql & "and serie ='" & Trim("" & my_serie) & "' " & Chr$(10)
            mysql = mysql & "AND numero ='" & Trim("" & my_numero) & "' " & Chr$(10)
            mysql = mysql & "AND tipo = '" & Trim("" & my_tipo) & "' " & Chr$(10)

        End If
   
        cn.Execute (mysql)
    Loop
    Close #fnum

    Exit Function
    
End Function

Public Function Enviar_Sunat(salida As Boolean, _
                             my_cantidad_file As Integer, _
                             my_caja As String)

    Dim mytable    As New ADODB.Recordset

    Dim mysql      As String

    Dim encontro   As Boolean

    Dim file       As String

    Dim input_file As String

    On Error GoTo Enviar_Sunat

    PathOrigen = "D:\ce_Input\FIRMADO\"

    sNextFile = Dir$("D:\ce_Input\FIRMADO\R_*.txt")

    While sNextFile <> ""

        conta = conta + 1
        ArchivoOrigen = PathOrigen & sNextFile
        file = sNextFile
        sNextFile = Dir$

        Do While encontro = False
            salida = FileExists("D:\ce_Input\FIRMADO\" & Left(file, (Len(file) - 0)))

            If salida = True Then
                input_file = "D:\ce_Input\FIRMADO\" & file
                encontro = True

            End If

            salida = FileExists("D:\ce_Input\ERROR\" & Left(file, (Len(file) - 0)))

            If salida = True Then
                input_file = "D:\ce_output\ERROR\" & file
                encontro = True

            End If

        Loop
        my_numero = Mid(file, 23, 8)
        my_t = Mid(file, 18, 1)

        If my_t = "F" Then
            my_tipo = "2"
        ElseIf my_t = "B" Then
            y_tipo = "1"

        End If

        my_serie = Mid(file, 19, 3)
        Call Graba_RSunat(input_file, my_serie, Val(my_numero), my_tipo, my_caja)

        frm_ESunat.ProgressBar1.Value = (conta / my_cantidad_file * 100)
        frm_ESunat.lbl_datos = "Elaborando Envio Sunat al.." & ((conta / my_cantidad_file) * 100) & "%"
    Wend
 
    Exit Function
Enviar_Sunat:

    Select Case Err.Number

        Case 380

            Resume Next

        Case Else

            Resume Next

    End Select

    Exit Function

End Function

'******10/02/2018 pll para las caja 1 persona varios ruc
Public Function read_caja(my_caja As String)

    Dim input_record As String

    Dim my_conta     As Integer

    Dim mytablef     As New ADODB.Recordset

    Dim mysql        As String

    Dim myDato       As String

    Dim myCuenta     As Integer

    input_file = globalpath & "\caja.txt"
    fnum = FreeFile
    Open input_file For Input As #fnum

    Do Until EOF(fnum)
    
        Line Input #fnum, input_record
 
        my_conta = my_conta + 1
    
        my_caja = "'" & Trim(Mid(input_record, 1, 2)) & "'"
    Loop
    Close #fnum

End Function

Public Function control_llegada_file(my_cantidad_file As Integer, my_caja As String)

    Dim mytable As New ADODB.Recordset

    Dim mysql   As String

    mysql = ""
    mysql = "SELECT * " & Chr$(10)
    mysql = mysql & "FROM FACTURA " & Chr$(10)
    mysql = mysql & "WHERE CDR <> ''" & Chr$(10)
    mysql = mysql & "and FECHA ='" & Format(Now, "dd/mm/yyyy") & "'" & Chr$(10)
    mysql = mysql & "and caja=" & my_caja & "" & Chr$(10)
    mysql = mysql & "and E_SUNAT ='N'" & Chr$(10)
 
    mytable.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytable.RecordCount > 0 Then  'si existe
        my_cantidad_file = mytable.RecordCount
    Else

    End If

    mytable.Close
    Exit Function

End Function

' Testing Proyecto Facturacion Electronica
Public Function verifica_estado_electronico(input_file As String)

    Dim input_record As String

    Dim my_conta     As Integer

    Dim mytablef     As New ADODB.Recordset

    fnum = FreeFile
    Open input_file For Input As #fnum

    Do Until EOF(fnum)
        Line Input #fnum, input_record
        my_conta = my_conta + 1

        If Mid(input_record, 3, 2) = "OK" Then
            salida = True
        Else
            salida = True

        End If
            
    Loop
    Close #fnum

    Exit Function

verifica_estado_electronico:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

    End Select
    
End Function

' Testing Proyecto Facturacion Electronica

' Testing Proyecto Facturacion Electronica 09/04/2018
Public Function verifica_respuestaResumen_Sunat(input_file As String, fecha As String)

    Dim input_record As String

    Dim my_conta     As Integer

    Dim mytablef     As New ADODB.Recordset

    Dim count        As Integer

    Dim rpta         As String

    Dim valor        As String

    Dim valor2       As String

    Dim valor3       As String

    Dim cant         As Integer

    Dim rptaDoc      As Integer

    Dim buscar       As String

    CantDoc = 0

    fnum = FreeFile

    Open input_file For Input As #fnum

    Do Until EOF(fnum)
        Line Input #fnum, input_record
        my_conta = my_conta + 1

        If Mid(input_record, 3, 2) = "OK" Then
            input_record = (Mid(input_record, 43, Len(input_record) - 28))
            Call Busca_Valores_Palotes(input_record)
            salida = True
        Else
            salida = False
            CantDoc = Obtiene_CantidadDocumentos(fecha)

            If CantDoc > 0 Then
                MsgBox "Falta RESUMEN DIARIO. Fecha: " & fecha, vbCritical

            End If

        End If
            
    Loop
    Close #fnum

    Exit Function

verifica_respuesta_Sunat:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

    End Select
    
End Function

' Testing Proyecto Facturacion Electronica 09/04/2018

Public Function Busca_Valores_Palotes(car As String) As String

    Dim I, sw, j As Integer

    Dim cant    As Integer

    Dim cantult As Integer

    Dim cadena  As String

    cadena = car
    cadfinal = ""

    Dim cadrpta     As String

    Dim tipodoc     As String

    Dim tipocorta   As String

    Dim seriedoc    As String

    Dim seriecorta  As String

    Dim numerodoc   As String

    Dim numerocorta As String

    Dim rptaDoc     As String

    cadrpta = ""
 
    tipodoc = ""
    seriedoc = ""
    numerodoc = ""
 
    rptaDoc = ""
 
    tipocorta = ""
    seriecorta = ""
    numerocorta = ""
 
    cant = 0
    cantult = 0

    I = 0
    j = 0
    
    Do While Len(cadena) > 0
        I = I + 1

        cadfinal = cadfinal & Mid(cadena, 1, 1)
        cadena = Right(car, Len(cadena) - 1)

        If Right(cadfinal, 1) = "|" Or Right(cadena, 1) = "" Then
            cant = cant + 1

        End If
        
        If cant = 1 Then
            If Len(tipodoc) = 0 Then
                tipodoc = Left(cadfinal, Len(cadfinal) - 1)
                cadfinal = ""

            End If

        End If
        
        If cant = 2 Then
            If Len(seriedoc) = 0 Then
                seriedoc = Left(cadfinal, Len(cadfinal) - 1)
                cadfinal = ""

            End If

        End If
        
        If cant = 3 Then
            If Len(numerodoc) = 0 Then
                numerodoc = Left(cadfinal, Len(cadfinal) - 1)
                cadfinal = ""

            End If

        End If
        
        If cant = 4 Then
            If Len(rptaDoc) = 0 Then
            
                For j = 1 To Len(cadfinal)

                    If Mid(cadfinal, j, 1) = "|" Then
                        cantult = cantult + 1

                    End If

                Next
                 
                If cantult = 0 Then
                    rptaDoc = Left(cadfinal, Len(cadfinal) - 1)
                    cadfinal = Right(cadfinal, 0)
                    cadena = cadfinal & cadena
                    cadfinal = ""
                Else
                    rptaDoc = Left(cadfinal, Len(cadfinal) - 4)
                    cadfinal = Right(cadfinal, 3)
                    cadena = cadfinal & cadena
                    cadfinal = ""

                End If
                
                If tipodoc = "03" Then ' Si tip es Boleta
                    my_tipo = "1"
                ElseIf tipodoc = "07" Then ' Si tipo es Nota Credito
                    my_tipo = "" & obtiene_TipoDocumentoNotas("E")
                ElseIf tipodoc = "08" Then ' Si tipo es Nota Débito
                    my_tipo = "" & obtiene_TipoDocumentoNotas("F")

                End If
                    
                Call Actualiza_Estado_Sunat(rptaDoc, my_local, my_tipo, seriedoc, numerodoc, my_acu)
                cadrpta = ""
                tipodoc = ""
                seriedoc = ""
                numerodoc = ""
                rptaDoc = ""
                     
                tipocorta = ""
                seriecorta = ""
                numerocorta = ""
                cant = 0
                cantult = 0
                ' End If
           
            End If

        End If
       
    Loop

End Function

' Testing Proyecto Facturacion Electronica 09/04/2018

' Testing Proyecto Facturacion Electronica 09/04/2018

' Testing Proyecto Facturacion Electronica
Public Sub Actualiza_Estado_Sunat(my_estado As String, _
                                  my_local As String, _
                                  my_tipo As String, _
                                  my_serie As String, _
                                  my_numero As String, _
                                  my_acu As String)
           
    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    mysql = ""
    mysql = "update  factura " & Chr$(10)
    mysql = mysql & " set ESTADO_SUNAT='" & "" & my_estado & "'" & Chr$(10)
    mysql = mysql & " where  local='" & my_local & "'" & Chr$(10)
    mysql = mysql & " and tipo='" & "" & my_tipo & "' " & Chr$(10)
    mysql = mysql & " and serie='" & "" & my_serie & "'" & Chr$(10)
    mysql = mysql & " and numero='" & "" & my_numero & "'" & Chr$(10)

    cn.Execute (mysql)

End Sub

' Testing Proyecto Facturacion Electronica

' Testing Proyecto Facturacion Electronica
Public Function verifica_estado_electronicoXDocumento(input_file As String)

    Dim input_record As String

    Dim my_conta     As Integer

    Dim mytablef     As New ADODB.Recordset

    fnum = FreeFile
    Open input_file For Input As #fnum

    Do Until EOF(fnum)
        Line Input #fnum, input_record
        my_conta = my_conta + 1

        If Mid(input_record, 3, 2) = "OK" Then
            input_record = (Mid(input_record, 8, Len(input_record)))
            my_tipo = Right(input_record, 2)
                
            Call Busca_Valores_ResumenXDocumento(input_record)
                
            salida = True
        Else
            salida = False

        End If
            
    Loop
    Close #fnum

    Exit Function

verifica_estado_electronico:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

    End Select
    
End Function

' Testing Proyecto Facturacion Electronica

Public Function Busca_Valores_ResumenXDocumento(car As String) As String

    Dim I, sw As Integer

    Dim cant    As Integer

    Dim cantult As Integer

    Dim cadena  As String

    cadena = car
    cadfinal = ""

    Dim cadrpta As String

    Dim rptaDoc As String

    cadrpta = ""
    rptaDoc = ""
 
    cant = 0

    I = 0
    
    Do While Len(cadena) > 0
        I = I + 1

        cadfinal = cadfinal & Mid(cadena, 1, 1)
        cadena = Right(car, Len(cadena) - 1)

        If Right(cadfinal, 1) = "|" Or Right(cadena, 1) = "" Then
            cant = cant + 1

        End If

        If cant = 1 Then
            If Len(rptaDoc) = 0 Then
                rptaDoc = Left(cadfinal, Len(cadfinal) - 1)
                cadena = ""
                cadfinal = ""

            End If
                
            If my_tipo = "01" Then
                my_tipo = "2"
            ElseIf my_tipo = "07" Or my_tipo = "08" Then
                my_tipo = obtiene_TipoDocumento(my_acu)

            End If

            Call Actualiza_Estado_Sunat(rptaDoc, my_local, my_tipo, my_serie, my_numero, my_acu)
            cadrpta = ""
            rptaDoc = ""
            cant = 0

        End If

    Loop
          
End Function

' Testing Proyecto Facturacion Electronica
Public Function obtiene_TipoDocumento(my_tipo As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select tipo from tipo where  TIPODOC='" & my_tipo & "' AND SUBSTRING(SERIE,1,1)='F' ", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        obtiene_TipoDocumento = mytablex.Fields("tipo")
    Else
        obtiene_TipoDocumento = "NO EXISTE"

    End If

    mytablex.Close

End Function

' Testing Proyecto Facturacion Electronica

' Testing Proyecto Facturacion Electronica
Public Function obtiene_TipoDocumentoNotas(my_tipo As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select tipo from tipo where  TIPODOC='" & my_tipo & "' AND SUBSTRING(SERIE,1,1)='B' ", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        obtiene_TipoDocumentoNotas = mytablex.Fields("tipo")
    Else
        obtiene_TipoDocumentoNotas = "NO EXISTE"

    End If

    mytablex.Close

End Function

' Testing Proyecto Facturacion Electronica

' Testing Proyecto Facturacion Electronica
Public Function Obtiene_CantidadDocumentos(fecha As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select count(numero) as cantidad from factura where estado='2' and  fecha='" & fecha & "' AND (acu='A' OR acu='B' OR acu='C' OR acu='D' or acu='E' or acu='F') ", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        Obtiene_CantidadDocumentos = mytablex.Fields("cantidad")

    End If

    mytablex.Close

End Function

' Testing Proyecto Facturacion Electronica

'Plan de Contingencia 07/05/2018
Function estrae_PlanContingencia(my_ruc As String, my_tipocont As String)

    Dim mysql       As String

    Dim mytable     As New ADODB.Recordset

    Dim hastaCuanto As Integer

    Dim nuevoDato   As String

    Dim myDato      As String

    mysql = ""
    mysql = "SELECT  *from factura " & Chr$(10)
    
    mytable.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytable.RecordCount > 0 Then
        salida = True
        hastaCuanto = 8 - Len(mytable.Fields("NUMERO")) '**en la tabla
        myDato = mytable.Fields("NUMERO")
        Call E_llenar_zero(hastaCuanto, myDato, my_numero)
        'PARA LA SERIE
 
        Call E_llenar_zero(4 - Len(mytable.Fields("SERIE")), mytable.Fields("SERIE"), my_serie)
     
        FileName = "D:\" & my_ruc & ".INPUT.TXT"
        file = my_ruc & ".INPUT.TXT"
        ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
 
        Do
            c = c + 1 'el contador sirve para la lista

            If mytable.EOF Then Exit Do
            'aqui es para abrir el file txt y llenarlo
            Filelibero1 = FreeFile
            Open FileName For Append As #Filelibero1
            conta = conta + 1
            conta_record = mytable.RecordCount

            If c = 1 Then

            End If
   
            myREG = ""
            myREG = myREG & "" & my_tipocont
            myREG = myREG & "|01|"
      
            hastaCuanto = 4 - Trim(Len(mytable.Fields("SERIE")))
            Call E_llenar_zero(3 - Len(mytable.Fields("SERIE")), mytable.Fields("SERIE"), nuevoDato)
            myREG = myREG & "" & Trim(nuevoDato)
            myREG = myREG & "-"
      
            If Trim(mytable.Fields("NUMERO")) = "vacio" Then
                myREG = myREG & "|" 'gion separador
            Else
                hastaCuanto = 8 - Trim$(Len(mytable.Fields("NUMERO"))) '**en la tabla
                myDato = mytable.Fields("NUMERO")
                Call E_llenar_zero(hastaCuanto, myDato, nuevoDato) 'NUMERO 3
                myREG = myREG & nuevoDato
                myREG = myREG & "|" 'gion separador

            End If
   
            If Trim(mytable.Fields("FECHA")) = 0 Then
                myREG = myREG & "|" 'gion separador
            Else
                Call llenar_datos(10 - Trim(Len(mytable.Fields("FECHA"))), mytable.Fields("FECHA"), nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "|" 'gion separador

            End If
   
            If Trim(mytable.Fields("tipo")) = "2" Then
                myREG = myREG & "01|"
            ElseIf Trim(mytable.Fields("tipo")) = "1" Then
                myREG = myREG & "03|"

            End If
   
            Print #Filelibero1, myREG
            Close #Filelibero1
            mytable.MoveNext
        Loop
        Close #Filelibero1

    End If

    mytable.Close

    'aqui copia el archivo txt

    origen = FileName
    Exit Function

estrae_PlanContingencia:
 
    Select Case Err.Number

        Case 71

            Resume Next

        Case 424

            Resume Next

        Case Else

            Resume Next

    End Select
 
End Function

'Plan de Contingencia 07/05/2018

