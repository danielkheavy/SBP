Attribute VB_Name = "S_NUBE_FACT"
'----------------------------------------------------------------------------------
'Variables para documento venta detalle provienen  al generar la venta
Global VarSerie             As String: Global VarNumer As String
Global VarNubeSerie As String
Global VarNubeNumero As String
Global VarRegistroCount As Integer
Dim i As Integer
Global crear_txt_para_enviar As String  ' Vamos a crear unas cuantas variables como ejemplo
Global crear_txt_para_enviarDetalle As String  ' Vamos a crear unas cuantas variables como ejemplo

Sub NUBEFACT_INSERT()
    
'    If varNOMBRE_CLIENTE = "" Then
'        varNOMBRE_CLIENTE = "-"
'
'    End If
'
'    If Len(VarClienteCodigo) = 11 Then
'        varTIPODOI = "6"
'        varDESC_TIPODOCU = "RUC"
'    ElseIf Len(VarClienteCodigo) = 8 Then
'        varTIPODOI = "1"
'        varDESC_TIPODOCU = "DNI"
'    ElseIf Len(VarClienteCodigo) = 12 Then
'        varTIPODOI = "4"
'        varDESC_TIPODOCU = "CARNET EXT."
'    ElseIf Len(VarClienteCodigo) = 9 Then
'        varTIPODOI = "7"
'        varDESC_TIPODOCU = "PASAPORTE"
'    ElseIf my_carga_busca_cliente(0).RUC = "" Then
'        varTIPODOI = "0"
'        varDESC_TIPODOCU = "OTROS"
'    ElseIf VarClienteCodigo <> "" Then
'        varTIPODOI = "0"
'        varDESC_TIPODOCU = "OTROS"
'
'    End If
'
'    '-------------------------------------------------------------
'    '-------------------------------------------------------------
'    'Solicitado por SENDA  si es boleta no registrar en blanco solo colocar un simbolo guion
'    If my_carga_busca_cliente(0).RUC = "" Then
'        varNUMERODOI = "-"
'    Else
'        varNUMERODOI = my_carga_busca_cliente(0).RUC
'
'    End If
'
'    If my_carga_busca_cliente(0).nombre = "" Then
'        varRAZONSOCIAL = "--"
'    Else
'        varRAZONSOCIAL = my_carga_busca_cliente(0).nombre
'
'    End If
'
'    If my_carga_busca_cliente(0).direccion = "" Then
'        varDIRECCIONFISCAL = "--"
'    Else
'        varDIRECCIONFISCAL = my_carga_busca_cliente(0).direccion
'
'    End If

    'Vamos a crear unas cuantas variables como ejemplo
'    crear_txt_para_enviar = "" & vbcrlf & _
'       "operacion|generar_comprobante|" & vbcrlf & _
'       "tipo_de_comprobante|" & varTIPODOCU & "|" & vbcrlf & _
'       "serie|" & VarNubeSerie & "|" & vbcrlf & _
'       "numero|" & VarNubeNumero & "|" & vbcrlf & _
'       "sunat_transaction|1|" & vbcrlf & _
'       "cliente_tipo_de_documento|" & Trim(varTIPODOI) & "|" & vbcrlf & _
'       "cliente_numero_de_documento|" & Trim(varCODIGO_CLIENTE) & "|" & vbcrlf & _
'       "cliente_denominacion|" & varNOMBRE_CLIENTE & "|" & vbcrlf & _
'       "cliente_direccion|" & varDIRECCIONFISCAL & "|" & vbcrlf & _
'       "cliente_email|" & varEMAIL_FROM & "|" & vbcrlf & _
'       "cliente_email_1||" & vbcrlf & _
'       "cliente_email_2||" & vbcrlf & _
'       "fecha_de_emision|" & Left(varFEMISION, 10) & "|" & vbcrlf & _
'       "fecha_de_vencimiento|" & Left(varFVENC, 10) & "|" & vbcrlf & _
'       "moneda|1|" & vbcrlf & _
'       "tipo_de_cambio||" & vbcrlf & _
'       "porcentaje_de_igv|18|" & vbcrlf & _
'       "descuento_global||" & vbcrlf & _
'       "total_descuento|" & varTOTAL_DSCTO & "|" & vbcrlf & _
'       "total_anticipo||" & vbcrlf & _
'       "total_gravada|" & varOP_GRAVADA & "|" & vbcrlf & _
'       "total_inafecta|" & varOP_INAFECTA & "|" & vbcrlf & _
'       "total_exonerada|" & varOP_EXONERADA & "|" & vbcrlf & _
'       "total_igv|" & varIGV & "|" & vbcrlf
'
'    crear_txt_para_enviar = crear_txt_para_enviar & _
'        "total_gratuita|" & varTOTAL_GRATUITAS & "|" & vbcrlf & _
'        "total_otros_cargos|" & varTOTAL_OTROS_CARGOS & "|" & vbcrlf & _
'        "total|" & varTOTAL_PAGAR & "|" & vbcrlf & _
'        "percepcion_tipo||" & vbcrlf & _
'        "percepcion_base_imponible||" & vbcrlf & _
'        "total_percepcion|" & varIMP_PERCEPCION & "|" & vbcrlf & _
'        "total_incluido_percepcion||" & vbcrlf & _
'        "detraccion|false|" & vbcrlf & _
'        "observaciones||" & vbcrlf & _
'        "enviar_automaticamente_a_la_sunat|true|" & vbcrlf & _
'        "enviar_automaticamente_al_cliente|false|" & vbcrlf & _
'        "codigo_unico||" & vbcrlf & _
'        "condiciones_de_pago||" & vbcrlf & _
'        "medio_de_pago||" & vbcrlf & _
'        "placa_vehiculo||" & vbcrlf & _
'        "orden_compra_servicio||" & vbcrlf & _
'        "tabla_personalizada_codigo||" & vbcrlf & _
'        "formato_de_pdf||" & vbcrlf



'"documento_que_se_modifica_tipo||" & vbcrlf & _
'        "documento_que_se_modifica_serie||" & vbcrlf & _
'        "documento_que_se_modifica_numero||" & vbcrlf & _
'        "tipo_de_nota_de_credito||" & vbcrlf & _
'        "tipo_de_nota_de_debito||" & vbcrlf & _

    'i = 1

   

    '" &  & "
    'crear_txt_para_enviar = crear_txt_para_enviar & _
    '"item|NIU|001|DETALLE DEL PRODUCTO|1|500|590||500|1|90|590|false|||" & vbcrlf & _
    '"item|ZZ|001|DETALLE DEL SERVICIO|5|20|23.60||100|1|18|118|false|||" & vbcrlf

    Dim TxtParaEnviar As String

    TxtParaEnviar = crear_txt_para_enviar
    Ruta = "https://www.pse.pe/api/v1/3db6b2c685b548c080c8abe558cc91c0023a2e6122e0488fabef3a7bb69e15b5"
    token = "eyJhbGciOiJIUzI1NiJ9.IjQ4NDdhYjAxZmQzMjRkYWE5N2IxNmU0NmJhYTIzMGY0Y2Y2ZThmNGY3NjNmNDM1OThhMGZiMWY1MDI3ODM2Nzki.czuZSlFNsuT4VY7i2JxXH9N3XUXtmZK_oORPhzWi98k"

    'MsgBox "Este es el archivo TXT que se enviará: " & TxtParaEnviar
    Dim myMSXML

    Dim variable As String

    Set XMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    XMLHTTP.Open "POST", Ruta, False
    XMLHTTP.setRequestHeader "Content-Type", "text/plain"
    XMLHTTP.setRequestHeader "Content-Length", Len(crear_txt_para_enviar)
    XMLHTTP.setRequestHeader "Authorization", "Token token=" & token
    XMLHTTP.send (crear_txt_para_enviar)

    'myMSXML.send crear_txt_para_enviar
    'myMSXML.send variable
    '''#########################################################
    ' #### PASO 4: LEER RESPUESTA DE NUBEFACT ####  ' +++++++++++++++++++++++++++++++++++++++++++++++++++++++
    ' # Recibirás una respuesta de NUBEFACT inmediatamente lo cual se debe leer, verificando que no haya errores.
    ' # Debes guardar en la base de datos la respuesta que te devolveremos.
    ' # Escríbenos a soporte@nubefact.com o llámanos al teléfono: 01 468 3535 (opción 2) o celular (WhatsApp) 955 598762
    ' # Puedes imprimir el PDF que nosotros generamos como también generar tu propia representación impresa previa coordinación con nosotros.
    ' # La impresión del documento seguirá haciéndose desde tu sistema. Enviaremos el documento por email a tu cliente si así lo indicas en el archivo JSON o TXT.
    ' +++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Dim TxtDeRespuesta As String

    Dim Explode()      As String

    TxtDeRespuesta = XMLHTTP.responseText
    
    'LEEMOS LA RESPUESTA EN TXT
    MsgBox (TxtDeRespuesta)
    
    'Text1.Text = TxtDeRespuesta
    Explode = Split(TxtDeRespuesta, "|")

    MsgBox (Explode(35) & "|" & Explode(36))
    'Text2.Text = Explode(35)
    'Text2.Text = Explode(36)
    'Explode = Split("tutorial, videotutorial", ",")
    
    'Call Buscar(TxtDeRespuesta)
    'TxtDeRespuesta.enlace_del_pdf
    'MsgBox TxtDeRespuesta
End Sub



