Attribute VB_Name = "SENDA_sp"
'danielkheavy 06/07/18
'FINALIDAD
'Integrar el Sistema de Facturación Electrónica SENDA EF@CT con Sistemas de Terceros.
'Para ello, se ha creado Procedimientos Almacenados, que servirán como interfaces para
'el registro y consulta de los comprobantes de pago electrónicos.
'PROCESO DE REGISTRO DE COMPROBANTE DE PAGO
'El Cliente, cuando emite o registre un comprobante de pago en su módulo de facturación, deberá ejecutar los Store Procedure en el siguiente orden:
'1. dbo. SENDAEFACT_INSERT_DOCUMENTOVENTA: Inserta la cabecera del comprobante de pago.
'2. dbo. SENDAEFACT_INSERT_DOCUMENTOVENTADETALLE: Inserta el detalle del  de pago.
'3. dbo. SENDAEFACT_FIRMAR_XML_ARG: Firma el Comprobante de pagcomprobanteo.
'4. dbo. SENDAEFACT_GET_DSVALUE: Obtiene el DSVALUE (Código Hash).

'----------------------------------------------------------------------------------
'Variables para documento venta detalle provienen  al generar la venta

Global varRUC_EMISOR             As String: Global varRAZONSOCIAL_EMISOR As String
Global varDIRECCIONFISCAL_EMISOR As String: Global varTELEFONO_EMISOR As String
Global varEMAIL_EMISOR           As String: Global varTIPOVENTASUNAT As String
Global varTIPODOCU               As String: Global varNRO_EFACT As String
Global varFEMISION               As Date:   Global varFVENC As Date
Global varTIPODOCUREFE           As String: Global varNUMEROREFE As String
Global varCODIGO_MOTIVO_07_08    As String: Global varDESCRIPCION_MOTIVO_07_08 As String
Global varTIPODOI                As String: Global varDESC_TIPODOCU As String
Global varNUMERODOI              As String: Global varRAZONSOCIAL As String
Global varDIRECCIONFISCAL        As String: Global varCODIGO_CLIENTE As String
Global varNOMBRE_CLIENTE         As String: Global varEMAIL_FROM As String
Global varEMAIL_CC               As String: Global varNOMB_VENDEDOR As String
Global varNOMB_RESERVA           As String: Global varFCHECKIN As Date
Global varNRO_RESERVA            As String
Global varFCHECKOUT              As Date: Global varNUMEROS_PREPAGOS As String
Global varMETODO_PAGO            As String: Global varPUNTO_EMISION As String
Global varCAJA                   As String: Global varNOMB_CAJERO As String
Global varSERIE_PRINT            As String: Global varNRO_TRANSACCCION As String
Global varORDEN_COMPRA           As String: Global varGLOSA As String
Global varGLOSA_DETRACCION       As String: Global varGLOSA_REFE As String
Global varMENSAJE                As String: Global varGLOSA_PIE_PAGINA As String
Global varTIPO_GR                As String: Global varNUMERO_GR As String
Global varOP_EXONERADA           As Double: Global varOP_INAFECTA As Double
Global varOP_GRAVADA             As Double: Global varISC As Double
Global varIGV                    As Double: Global varIMP_PERCEPCION As Double
Global varIMP_TOTAL              As Double: Global varTOTAL_PAGAR As Double
Global varTOTAL_OTROS_TRIBUTOS   As Double: Global varTOTAL_OTROS_CARGOS As Double
Global varTOTAL_DSCTO            As Double: Global varPORC_DSCTOGLOBAL As Double
Global varDSCTO_GLOBAL           As Double: Global varTOTAL_GRATUITAS As Double
Global varTOTAL_ANTICIPO         As Double: Global varTIPO_MONEDA As String
Global varIMPORTE_LETRAS         As String: Global varTOTAL_SERVICIO As Double
Global varPORC_SERVICIO          As Double: Global varPROPINA As Double:
Global varPRESTAMO               As Double: Global varVUELTO  As Double
Global varCON_DETALLE            As String: Global varUSUARIO As String
Global varPORC_IGV               As Double: Global varTIPOCAMBIO As Double
Global varSUCU As String


'--------'Otras Variables para otros sub--------------------------------
Global varSTR_DS_VALUE            As String:
Global varNrolocal As String
Global RePrintRUC_EMISOR As String
Global RePrintvarNRO_EFACT As String
Global RePrintvarTIPODOCU As String

'------------------------------------------------
'Variables que llenaran el detalle de la factura documento venta detalle
Global dvarINDICE                As Integer
Global dvaris_anticipo           As Integer
Global dvarCODIGO                As String
Global dvarNOMB_ITEM             As String
Global dvarCANTIDAD              As Double
Global dvarUM                    As String
Global dvarVALORUNITBRUTO        As Double
Global dvarVALOR_UNIT            As Double
Global dvarDSCTO_UNIT            As Double
Global dvarPORC_DESCTO           As Double
Global dvarIGV_UNIT              As Double
Global dvarISC                   As Double
Global dvarPRECIO_UNIT_BRUTO     As Double
Global dvarPRECIO_UNIT           As Double
Global dvarVALORVENTABRUTO       As Double
Global dvarVALOR_VENTA           As Double
Global dvarPRECIO_VENTA          As Double
Global dvarTIPO_AFECT_IGV        As String
Global dvarCODIGO_TRIBUTO        As String
Global dvarTIPO_PRECIOVENTA      As String
Global dvarVISIBLE               As Integer
'-------------------------------------------------------------------------------
'Variables de Documento de Baja
Global bvarFBAJA                 As String
Global bvarMOTIVO_BAJA           As String

Sub SENDAEFACT_INSERT_DOCUMENTOVENTA()
    'Asignacion valores por defecto de factura o documento de venta
    varCON_DETALLE = "S"

    ' Variables Capturador de Error
    Dim varSTR_ERROR As String
    Dim varCOD_ERROR As Integer
    
    Dim objConn      As ADODB.Connection 'variable Objeto de conexion a la BD
    Dim objCmd       As ADODB.Command    'variable comando de ejecucion SP (Store Procedure)
    'Dim prmParam As ADODB.Parameter
    ' prmParam = 18
    '-------------------
    '--- Abrir Coneccion con la BD
    '-------------------
    Set objConn = New ADODB.Connection
    objConn.Provider = "SQLOLEDB"
    'Conexion sencilla de prueba
    'objConn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; pwd=mastercard;Initial Catalog=daniel;Data Source=VISITEC-PC"
    
    'Conexion con parametro Output y BD SENDA
    objConn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; pwd=mastercard;Initial Catalog=SENDAEFACT_LOCAL;Data Source=(LOCAL)"
    objConn.Open ' abrimos nuestra conexion
    '-------------------
    '--- Asignamos el Objeto comando para ejecutar store procedure
    '-------------------
    Set objCmd = New ADODB.Command
    objCmd.CommandType = adCmdStoredProc ' indicamos que es de tipo SP
    
    'objCmd.CommandText = "InsertarUsuario" aqui para BD de prueba daniel
    objCmd.CommandText = "SENDAEFACT_INSERT_DOCUMENTOVENTA"    ' BD SENDA aca debemos indicar el nombre de SP que se va a tratar
    Set objCmd.ActiveConnection = objConn 'asignamo la conexion a la BD al comando
     
    '-------------------
    '--- Creamos y asignamos los parametros
    '-------------------
    'En esta Seccion se debera indicar cada uno de los parametro segun el requerimiento de cada SP
    'DECLARE @RC int
    
    objCmd.Parameters.Append objCmd.CreateParameter("@RUC_EMISOR", adChar, adParamInput, 11, varRUC_EMISOR) 'RUC EMISOR
    objCmd.Parameters.Append objCmd.CreateParameter("@RAZONSOCIAL_EMISOR", adVarChar, adParamInput, 250, varRAZONSOCIAL_EMISOR) 'RAZON DEL EMISOR
    objCmd.Parameters.Append objCmd.CreateParameter("@DIRECCIONFISCAL_EMISOR", adVarChar, adParamInput, 250, varDIRECCIONFISCAL_EMISOR) 'DIRECCION DEL EMISOR
    objCmd.Parameters.Append objCmd.CreateParameter("@TELEFONO_EMISOR", adVarChar, adParamInput, 100, varTELEFONO_EMISOR) 'TELEFONO DEL EMISOR
    objCmd.Parameters.Append objCmd.CreateParameter("@EMAIL_EMISOR", adVarChar, adParamInput, 100, varEMAIL_EMISOR) 'CORREO DEL EMISOR
    objCmd.Parameters.Append objCmd.CreateParameter("@TIPOVENTASUNAT", adChar, adParamInput, 2, varTIPOVENTASUNAT) 'SI EL TIPO DE OPERACION ES 01= VENTA INTERNA, SI ES 02=VENTA EXPORTACION, SI ES 03= NO DOMICILIADO
    objCmd.Parameters.Append objCmd.CreateParameter("@IS_GRATUITA", adInteger, adParamInput, 2, 0) 'GRATUITA, VALOR POR DEFECTO CERO.
    objCmd.Parameters.Append objCmd.CreateParameter("@TIPODOCU", adChar, adParamInput, 2, varTIPODOCU)  'TIPO DE COMPROBANTE DE PAGO (01,03,07,08)
    objCmd.Parameters.Append objCmd.CreateParameter("@NRO_EFACT", adVarChar, adParamInput, 12, varNRO_EFACT) ' NUMERO DE COMPROBANTE DE PAGO EJEM: F00100000001 (SERIE =4 CARACTERES)(CORRELATIVO 8 CARACTERES)
    objCmd.Parameters.Append objCmd.CreateParameter("@FEMISION", adDate, adParamInput, 10, varFEMISION) 'FECHA DE EMISION DEL DOCUMENTO formato año-mes-dia
    objCmd.Parameters.Append objCmd.CreateParameter("@FVENC", adDate, adParamInput, 10, varFVENC) 'FECHA DE VENCIMIENTO DEL DOCUMENTO
    objCmd.Parameters.Append objCmd.CreateParameter("@TIPODOCUREFE", adChar, adParamInput, 2, varTIPODOCUREFE) 'TIPO DE DOCUMENTO DE REFERENCIA CUANDO ES UNA NOTA DE CREDITO O NOTA DE DEBITO
    objCmd.Parameters.Append objCmd.CreateParameter("@NUMEROREFE", adVarChar, adParamInput, 50, varNUMEROREFE) 'NUMERO DE DOCUMENTO DE REFERENCIA CUANDO ES UNA NOTA DE CREDITO O NOTA DE DEBITO
    objCmd.Parameters.Append objCmd.CreateParameter("@CODIGO_MOTIVO_07_08", adChar, adParamInput, 2, varCODIGO_MOTIVO_07_08) 'CODIGO DEL MOTIVO CUANDO ES UNA NOTA DE CREDITO O NOTA DE DEBITO SEGÚN CATALOGO NRO 09 Y 10 SUNAT
    objCmd.Parameters.Append objCmd.CreateParameter("@DESCRIPCION_MOTIVO_07_08", adVarChar, adParamInput, 100, varDESCRIPCION_MOTIVO_07_08) 'DESCRIPCION DEL MOTIVO CUANDO ES UNA NOTA DE CREDITO O NOTA DE DEBITO
    objCmd.Parameters.Append objCmd.CreateParameter("@TIPODOI", adChar, adParamInput, 1, varTIPODOI) 'TIPO DE DOCUMENTO DE IDENTIDAD CATALOGO NRO 06 1=DNI, 0=Otro tipo de documento, 4= carnet Extranjeria
    objCmd.Parameters.Append objCmd.CreateParameter("@DESC_TIPODOCU", adVarChar, adParamInput, 50, varDESC_TIPODOCU) ' DESCRIPCION DEL TIPO DE DOCUMENTO DE IDENTIDAD DEL USUARIO POR EJEMPLO: DNI , RUC, ETC.
    objCmd.Parameters.Append objCmd.CreateParameter("@NUMERODOI", adVarChar, adParamInput, 20, varNUMERODOI) 'NUMERO DE DOCUMENTO DE INDENTIDAD
    objCmd.Parameters.Append objCmd.CreateParameter("@RAZONSOCIAL", adVarChar, adParamInput, 250, varRAZONSOCIAL) 'RAZON SOCIAL DEL ADQUIRIENTE O USUARIO
    objCmd.Parameters.Append objCmd.CreateParameter("@DIRECCIONFISCAL", adVarChar, adParamInput, 250, varDIRECCIONFISCAL) 'DIRECCION FISCAL SOCIAL DEL ADQUIRIENTE O USUARIO
    objCmd.Parameters.Append objCmd.CreateParameter("@CODIGO_CLIENTE", adVarChar, adParamInput, 50, varCODIGO_CLIENTE) 'CODIGO INTERNO DEL CLIENTE
    objCmd.Parameters.Append objCmd.CreateParameter("@NOMBRE_CLIENTE", adVarChar, adParamInput, 250, varNOMBRE_CLIENTE) 'NOMBRE DEL CLIENTE
    objCmd.Parameters.Append objCmd.CreateParameter("@EMAIL_FROM", adVarChar, adParamInput, 250, varEMAIL_FROM) 'EMAIL DEL CLIENTE SI MAS DE 1 EMAIL SEPARAR CON PUNTO Y COMA.
    objCmd.Parameters.Append objCmd.CreateParameter("@EMAIL_CC", adVarChar, adParamInput, 250, varEMAIL_CC) 'EMAIL CON COPIA SI MAS DE 1 EMAIL SEPARAR CON PUNTO Y COMA.
    objCmd.Parameters.Append objCmd.CreateParameter("@NOMB_VENDEDOR", adVarChar, adParamInput, 100, varNOMB_VENDEDOR) 'NOMBRE DEL VENDEDOR
    objCmd.Parameters.Append objCmd.CreateParameter("@NOMB_RESERVA", adVarChar, adParamInput, 50, varNOMB_RESERVA) 'NOMBRE DE LA RESERVA
    objCmd.Parameters.Append objCmd.CreateParameter("@NRO_RESERVA", adVarChar, adParamInput, 50, varNRO_RESERVA) 'NUMERO DE RESERVA
    objCmd.Parameters.Append objCmd.CreateParameter("@FCHECKIN", adDBDate, adParamInput, 10, varFCHECKIN) 'FECHA DE CHECK-IN
    objCmd.Parameters.Append objCmd.CreateParameter("@FCHECKOUT", adDate, adParamInput, 10, varFCHECKOUT) 'FECHA DE CHECK-OUT
    objCmd.Parameters.Append objCmd.CreateParameter("@NUMEROS_PREPAGOS ", adVarChar, adParamInput, 150, varNUMEROS_PREPAGOS) 'NUMEROS DE DOCUMENTO DE PREPAGOS
    objCmd.Parameters.Append objCmd.CreateParameter("@METODO_PAGO", adVarChar, adParamInput, 50, varMETODO_PAGO) 'CONDICION DE PAGO POR EJEMPLO: AL CONTADO O CREDITO
    objCmd.Parameters.Append objCmd.CreateParameter("@PUNTO_EMISION", adVarChar, adParamInput, 250, varPUNTO_EMISION) 'PUNTO O DIRECCION DE EMISION DEL COMPROBANTE DE PAGO
    objCmd.Parameters.Append objCmd.CreateParameter("@CAJA", adVarChar, adParamInput, 50, varCAJA) 'NUMERO O DESCRIPCION DE LA CAJA
    objCmd.Parameters.Append objCmd.CreateParameter("@NOMB_CAJERO", adVarChar, adParamInput, 50, varNOMB_CAJERO) 'NUMERO O DESCRIPCION DE LA CAJA
    objCmd.Parameters.Append objCmd.CreateParameter("@SERIE_PRINT", adVarChar, adParamInput, 50, varSERIE_PRINT) 'NUMERO DE LA SERIE DE LA IMPRESORA
    objCmd.Parameters.Append objCmd.CreateParameter("@NRO_TRANSACCCION", adVarChar, adParamInput, 50, varNRO_TRANSACCCION) 'NUMERO DE TRANSACCION
    objCmd.Parameters.Append objCmd.CreateParameter("@ORDEN_COMPRA", adVarChar, adParamInput, 50, varORDEN_COMPRA) 'NUMERO DE LA ORDEN DE COMPRA
    objCmd.Parameters.Append objCmd.CreateParameter("@GLOSA", adVarChar, adParamInput, 250, varGLOSA) 'GLOSA DEL COMPROBANTE DE PAGO
    objCmd.Parameters.Append objCmd.CreateParameter("@GLOSA_DETRACCION", adVarChar, adParamInput, 250, varGLOSA_DETRACCION) 'GLOSA PARA ESPECIFICAR LA DESCRIPCION DEL NUMERO DE LA RESOLUCION CUANDO ES AGENTE DE RETENCION LA CUAL SERVIRA PARA VISUALIZARLO EN EL COMPROBANTE DE PAGO
    objCmd.Parameters.Append objCmd.CreateParameter("@GLOSA_REFE", adVarChar, adParamInput, 250, varGLOSA_REFE) 'GLOSA DE REFERENCIA PARA VISUALIZAR EN EL COMPROBANTE DE PAGO
    objCmd.Parameters.Append objCmd.CreateParameter("@MENSAJE", adVarChar, adParamInput, 250, varMENSAJE) 'LEYENDA DEL COMPROBANTE DE PAGO EJEMPLO:TRANSFERENCIA GRATUITA DE BIENES Y SERVICIOS
    objCmd.Parameters.Append objCmd.CreateParameter("@GLOSA_PIE_PAGINA", adVarChar, adParamInput, 250, varGLOSA_PIE_PAGINA) 'GLOSA PARA EL PIE DE PAGINA PARA VISUAIZAR EN EL COMPROBANTE DE PAGO
    objCmd.Parameters.Append objCmd.CreateParameter("@TIPO_GR", adChar, adParamInput, 2, varTIPO_GR) 'TIPO DE DOCUMENTO DE LA GUIA DE REMISION RELACIONADA CON LA OPERACION QUE SE FACTURE
    objCmd.Parameters.Append objCmd.CreateParameter("@NUMERO_GR", adVarChar, adParamInput, 50, varNUMERO_GR) 'NUMERO DE OTRO DOCUMENTO RELACIONADO CON LA OPERACIÓN QUE SE FACTURE
    objCmd.Parameters.Append objCmd.CreateParameter("@OP_EXONERADA", adDouble, adParamInput, 9, varOP_EXONERADA) 'TOTAL DE LAS OPERACION EXONERADAS
    objCmd.Parameters.Append objCmd.CreateParameter("@OP_INAFECTA", adDouble, adParamInput, 9, varOP_INAFECTA) 'TOTAL DE LAS OPERACION INAFECTAS
    objCmd.Parameters.Append objCmd.CreateParameter("@OP_GRAVADA", adDouble, adParamInput, 9, varOP_GRAVADA)  'TOTAL DE LAS OPERACION GRAVADA
    objCmd.Parameters.Append objCmd.CreateParameter("@ISC", adDouble, adParamInput, 18, varISC) 'IMPUESTO SELECTIVO AL CONSUMO
    objCmd.Parameters.Append objCmd.CreateParameter("@IGV", adDouble, adParamInput, 18, varIGV) 'TOTAL DE IMPUESTO O IVA
    objCmd.Parameters.Append objCmd.CreateParameter("@IMP_PERCEPCION", adDouble, adParamInput, 18, varIMP_PERCEPCION) 'TOTAL EN SOLES DE LA PERCEPCION
    objCmd.Parameters.Append objCmd.CreateParameter("@IMP_TOTAL", adDouble, adParamInput, 18, varIMP_TOTAL) 'IMPORTE TOTAL
    objCmd.Parameters.Append objCmd.CreateParameter("@TOTAL_PAGAR", adDouble, adParamInput, 18, varTOTAL_PAGAR) 'TOTAL A PAGAR
    objCmd.Parameters.Append objCmd.CreateParameter("@TOTAL_OTROS_TRIBUTOS", adDouble, adParamInput, 18, varTOTAL_OTROS_TRIBUTOS) 'TOTAL DE OTROS TRIBUTOS
    objCmd.Parameters.Append objCmd.CreateParameter("@TOTAL_OTROS_CARGOS", adDouble, adParamInput, 18, varTOTAL_OTROS_CARGOS) 'TOTAL DE OTROS CARGOS
    objCmd.Parameters.Append objCmd.CreateParameter("@TOTAL_DSCTO", adDouble, adParamInput, 18, varTOTAL_DSCTO) 'TOTAL DE DESCUENTOS DEL DETALLE
    objCmd.Parameters.Append objCmd.CreateParameter("@PORC_DSCTOGLOBAL", adDouble, adParamInput, 18, varPORC_DSCTOGLOBAL) 'PORCENTAJE APLICADO AL DESTO GLOBAL DEL DOCUMENTO
    objCmd.Parameters.Append objCmd.CreateParameter("@DSCTO_GLOBAL", adDouble, adParamInput, 18, varDSCTO_GLOBAL) 'TOTAL DE DESCUENTO GLOBAL DEL DOCUMENTO
    objCmd.Parameters.Append objCmd.CreateParameter("@TOTAL_GRATUITAS", adDouble, adParamInput, 18, varTOTAL_GRATUITAS) 'TOTAL DE OPERACIONES GRATUITAS
    objCmd.Parameters.Append objCmd.CreateParameter("@TOTAL_ANTICIPO", adDouble, adParamInput, 18, varTOTAL_ANTICIPO) 'TOTAL DE ANTICIPO
    objCmd.Parameters.Append objCmd.CreateParameter("@TIPO_MONEDA", adVarChar, adParamInput, 3, varTIPO_MONEDA)  'PEN = SOL Y USD =DOLARES AMERICANOS
    objCmd.Parameters.Append objCmd.CreateParameter("@IMPORTE_LETRAS", adVarChar, adParamInput, 250, varIMPORTE_LETRAS) 'IMPORTE EN LETRAS EJEMPLO:NUEVE CIENTOS ONCE CON 60/100 SOLES
    objCmd.Parameters.Append objCmd.CreateParameter("@TOTAL_SERVICIO", adDouble, adParamInput, 18, varTOTAL_SERVICIO) 'TOTAL DEL SERVICIO PARA HOTELES
    objCmd.Parameters.Append objCmd.CreateParameter("@PORC_SERVICIO", adDouble, adParamInput, 18, varPORC_SERVICIO) 'PORC DEL SERVICIO PARA HOTELES
    objCmd.Parameters.Append objCmd.CreateParameter("@PROPINA", adDouble, adParamInput, 18, varPROPINA) 'IMPORTE DE PROPINA PARA REST Y HOTELES. SI EXISTE VALOR DEBE IR EN LA SUMATORIA DEL CAMPO
    objCmd.Parameters.Append objCmd.CreateParameter("@PRESTAMO", adDouble, adParamInput, 18, varPRESTAMO) 'IMPORTE DE PRESTAMO PARA REST Y HOTELES.SI EXISTE VALOR DEBE IR EN LA SUMATORIA DEL CAMPOR
    objCmd.Parameters.Append objCmd.CreateParameter("@VUELTO", adDouble, adParamInput, 18, varVUELTO) 'IMPORTE DE VUELTO POR EL PAGO DEL DOCUMENTO
    objCmd.Parameters.Append objCmd.CreateParameter("@CON_DETALLE", adChar, adParamInput, 1, varCON_DETALLE)  'SI EL DOCUMENTO TIENE DETALLE POR DEFAULT = S
    objCmd.Parameters.Append objCmd.CreateParameter("@USUARIO", adVarChar, adParamInput, 50, varUSUARIO) 'NOMBRE DEL USUARIO DEL SISTEMA QUE CREO EL COMPROBANTE DE PAGO
    objCmd.Parameters.Append objCmd.CreateParameter("@PORC_IGV", adDouble, adParamInput, 18, varPORC_IGV) '% DE IGV EJEMPO:18
    objCmd.Parameters.Append objCmd.CreateParameter("@TIPOCAMBIO", adDouble, adParamInput, 18, varTIPOCAMBIO) 'TIPO DE CAMBIO VENTA
    objCmd.Parameters.Append objCmd.CreateParameter("@SUCU", adChar, adParamInput, 2, varSUCU) 'Sucursal
    
    'objCmd.Parameters.Append objCmd.CreateParameter("@STR_ERROR", adVarChar, adParamOutput, 50, "")
    'objCmd.Parameters.Append objCmd.CreateParameter("@COD_ERROR", adInteger, adParamOutput, 3, 0)
    
    objCmd.Parameters.Append objCmd.CreateParameter("STR_ERROR", adVarChar, adParamOutput, 50, InCadena)
    objCmd.Parameters.Append objCmd.CreateParameter("COD_ERROR", adInteger, adParamOutput, 3, InNumero)
    
    'objCmd.Parameters.Append objCmd.CreateParameter("@COD_ERROR", adInteger, adParamOutput, 2, varCOD_ERROR)
    
    'objCmd.Parameters.Append objCmd.CreateParameter("@STR_ERROR", adVarChar, adParamOutput, 50, "12456")
    'objCmd.Parameters.Append objCmd.CreateParameter("@COD_ERROR", adInteger, adParamOutput, 50, "12456")
    
    '-------------------
    '--- Ejecutamos el Store Procedure
    '-------------------
    objCmd.Execute
    '-------------------
    
    'txtUserName.Text = Trim(objCmd.Parameters(70).Value)
    'txtClave.Text = objCmd.Parameters(71).Value
    
    'Para el ejemplo capturamos la salida o respuesta output que devuelve el SP y capturamos
    'su valor segun la matrix de variables y lo mostramos en unos Txtbox de Ejemplo pudiendo usarse para guardar en una BD
     
    varSTR_ERROR = objCmd.Parameters(71).Value
    varCOD_ERROR = objCmd.Parameters(72).Value
    
    'FrmPrincipal.txtusername.Text = varSTR_ERROR
    'FrmPrincipal.txtClave.Text = varCOD_ERROR
    '-------------------
    
    '-------------------
    '----- Le asignamos nada al comando y conexion y se cierra
    objConn.Close
    Set objCmd = Nothing
    Set objConn = Nothing
    '-------------
    'MsgBox ("Documento Venta Insertado"), vbOKOnly, "Inser_DocumentoVenta"

End Sub

Sub SENDAEFACT_INSERT_DOCUMENTOVENTADETALLE()
    ' Variables Capturador de Error
    Dim varSTR_ERROR As String
    Dim varCOD_ERROR As Integer
    
    Dim objConn      As ADODB.Connection 'variable Objeto de conexion a la BD
    Dim objCmd       As ADODB.Command    'variable comando de ejecucion SP (Store Procedure)
    '-------------------
    '--- Abrir Coneccion con la BD
    '-------------------
    Set objConn = New ADODB.Connection
    objConn.Provider = "SQLOLEDB"
    'Conexion sencilla de prueba
    'objConn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; pwd=mastercard;Initial Catalog=daniel;Data Source=VISITEC-PC"
    
    'Conexion con parametro Output y BD SENDA
    objConn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; pwd=mastercard;Initial Catalog=SENDAEFACT_LOCAL;Data Source=(LOCAL)"
    objConn.Open ' abrimos nuestra conexion
    '-------------------
    '--- Asignamos el Objeto comando para ejecutar store procedure
    '-------------------
    Set objCmd = New ADODB.Command
    objCmd.CommandType = adCmdStoredProc ' indicamos que es de tipo SP
    
    'objCmd.CommandText = "InsertarUsuario" aqui para BD de prueba daniel
    objCmd.CommandText = "SENDAEFACT_INSERT_DOCUMENTOVENTADETALLE"    ' BD SENDA aca debemos indicar el nombre de SP que se va a tratar
    Set objCmd.ActiveConnection = objConn 'asignamo la conexion a la BD al comando
     
    '-------------------
    '--- Creamos y asignamos los parametros
    '-------------------
    'En esta Seccion se debera indicar cada uno de los parametro segun el requerimiento de cada SP
    'DECLARE @RC int

    objCmd.Parameters.Append objCmd.CreateParameter("@RUC_EMISOR", adChar, adParamInput, 11, varRUC_EMISOR) 'RUC EMISOR
    objCmd.Parameters.Append objCmd.CreateParameter("@TIPODOCU", adChar, adParamInput, 2, varTIPODOCU) 'TIPO DE COMPROBANTE DE PAGO (01,03,07,08)
    objCmd.Parameters.Append objCmd.CreateParameter("@NRO_EFACT", adVarChar, adParamInput, 12, varNRO_EFACT) ' NUMERO DE COMPROBANTE DE PAGO EJEM: F00100000001 (SERIE =4 CARACTERES)(CORRELATIVO 8 CARACTERES)
    objCmd.Parameters.Append objCmd.CreateParameter("@INDICE", adInteger, adParamInput, 4, dvarINDICE) 'INDICE DEL ITEM
    objCmd.Parameters.Append objCmd.CreateParameter("@is_anticipo", adInteger, adParamInput, 1, dvaris_anticipo) 'SI EL ITEM ES UN ANTICIPO
    objCmd.Parameters.Append objCmd.CreateParameter("@CODIGO", adVarChar, adParamInput, 50, dvarCODIGO) 'CODIGO DEL ITEM O SERVICIO
    objCmd.Parameters.Append objCmd.CreateParameter("@NOMB_ITEM", adVarChar, adParamInput, 250, dvarNOMB_ITEM) 'DESCRIPCION DE ITEM O SERVICIO
    objCmd.Parameters.Append objCmd.CreateParameter("@CANTIDAD", adDouble, adParamInput, 18, dvarCANTIDAD) 'CANTIDAD DE ITEM O SERVICIO
    objCmd.Parameters.Append objCmd.CreateParameter("@UM", adVarChar, adParamInput, 10, dvarUM)  'CODIGO DE UNIDAD DE MEDIDA SEGUN TABLA SUNAT
    objCmd.Parameters.Append objCmd.CreateParameter("@VALORUNITBRUTO", adDouble, adParamInput, 18, dvarVALORUNITBRUTO) ' VALOR UNITARIO BRUTO SIN DESCUENTO DEL ITEM O SERVICIO
    objCmd.Parameters.Append objCmd.CreateParameter("@VALOR_UNIT", adDouble, adParamInput, 18, dvarVALOR_UNIT) ' VALOR UNITARIO DEL ITEM O SERVICIO
    objCmd.Parameters.Append objCmd.CreateParameter("@DSCTO_UNIT", adDouble, adParamInput, 18, dvarDSCTO_UNIT) ' DESCUENTO UNITARIO
    objCmd.Parameters.Append objCmd.CreateParameter("@PORC_DESCTO", adDouble, adParamInput, 18, dvarPORC_DESCTO) ' PORCENTAJE DE DESCTO
    objCmd.Parameters.Append objCmd.CreateParameter("@IGV_UNIT", adDouble, adParamInput, 18, dvarIGV_UNIT) ' IGV UNITARIO DEL ITEM O SERVICIO
    objCmd.Parameters.Append objCmd.CreateParameter("@ISC", adDouble, adParamInput, 18, dvarISC) ' IGV UNITARIO DEL ITEM O SERVICIO
    objCmd.Parameters.Append objCmd.CreateParameter("@PRECIO_UNIT_BRUTO", adDouble, adParamInput, 18, dvarPRECIO_UNIT_BRUTO) ' PRECIO BRUTO UNITARIO DEL ITEM O SERVICIO
    objCmd.Parameters.Append objCmd.CreateParameter("@PRECIO_UNIT", adDouble, adParamInput, 18, dvarPRECIO_UNIT) ' PRECIO UNITARIO DEL ITEM O SERVICIO
    objCmd.Parameters.Append objCmd.CreateParameter("@VALORVENTABRUTO", adDouble, adParamInput, 18, dvarVALORVENTABRUTO) ' VENTA BRUTO = @VALORUNITBRUTO * CANTIDAD
    objCmd.Parameters.Append objCmd.CreateParameter("@VALOR_VENTA", adDouble, adParamInput, 18, dvarVALOR_VENTA) ' EL VALOR VENTA =@VALOR_UNIT* @CANTIDAD
    objCmd.Parameters.Append objCmd.CreateParameter("@PRECIO_VENTA", adDouble, adParamInput, 18, dvarPRECIO_VENTA) ' EL PRECIO VENTA =@PRECIO_UNIT* @CANTIDAD
    objCmd.Parameters.Append objCmd.CreateParameter("@TIPO_AFECT_IGV", adChar, adParamInput, 2, dvarTIPO_AFECT_IGV) 'TIPO DE AFECTACION DEL IGV SEGUN CATALOGO NRO 7 SUNAT
    objCmd.Parameters.Append objCmd.CreateParameter("@CODIGO_TRIBUTO", adChar, adParamInput, 4, dvarCODIGO_TRIBUTO) 'CODIGO DE TRIBUTO DEGUN CATALOGO NRO 5 SUNAT
    objCmd.Parameters.Append objCmd.CreateParameter("@TIPO_PRECIOVENTA", adChar, adParamInput, 2, dvarTIPO_PRECIOVENTA) 'TIPO DE PRECIO DE VENTA SEGUN CATALOGO NRO 16 SUNAT
    objCmd.Parameters.Append objCmd.CreateParameter("@VISIBLE", adInteger, adParamInput, 1, dvarVISIBLE) 'SI ITEM SERA VISIBLE EN EL FORMATO DE IMPRESIÓN
        
    objCmd.Parameters.Append objCmd.CreateParameter("STR_ERROR", adVarChar, adParamOutput, 50, InCadena) 'DESCRIPCION DEL ERROR
    objCmd.Parameters.Append objCmd.CreateParameter("COD_ERROR", adInteger, adParamOutput, 3, InNumero) 'SI DEVUELTO -1 ES UN ERROR
       
    '-------------------
    '--- Ejecutamos el Store Procedure
    '-------------------
    objCmd.Execute
    '-------------------
    
    'txtUserName.Text = Trim(objCmd.Parameters(70).Value)
    'txtClave.Text = objCmd.Parameters(71).Value
    
    'Para el ejemplo capturamos la salida o respuesta output que devuelve el SP y capturamos
    'su valor segun la matrix de variables y lo mostramos en unos Txtbox de Ejemplo pudiendo usarse para guardar en una BD
    If IsNull(objCmd.Parameters(24).Value) Then
        varSTR_ERROR = "String Error Valor Nulo pero exitoso"
        'MsgBox (varSTR_ERROR), vbOKOnly, "Inser_DocumentoVenta"
    Else
        varSTR_ERROR = objCmd.Parameters(24).Value
    End If

    If IsNull(objCmd.Parameters(25).Value) Then
        varCOD_ERROR = -53
    Else
        varCOD_ERROR = objCmd.Parameters(25).Value
    End If
    
    'FrmPrincipal.txtusername.Text = varSTR_ERROR
    'FrmPrincipal.txtClave.Text = varCOD_ERROR
    '-------------------
    
    '-------------------
    '----- Le asignamos nada al comando y conexion CERRAMOS
    objConn.Close
    Set objCmd = Nothing
    Set objConn = Nothing
    '-------------
    'MsgBox ("Documento Venta Detalle"), vbOKOnly, "Inser_DocumentoVenta"
End Sub

Sub SENDAEFACT_FIRMAR_XML_ARG()
    ' Variables Capturador de Error
    Dim varSTR_ERROR As String
    Dim varCOD_ERROR As Integer
    
    Dim objConn      As ADODB.Connection 'variable Objeto de conexion a la BD
    Dim objCmd       As ADODB.Command    'variable comando de ejecucion SP (Store Procedure)
    '-------------------
    '--- Abrir Coneccion con la BD
    '-------------------
    Set objConn = New ADODB.Connection
    objConn.Provider = "SQLOLEDB"
    'Conexion sencilla de prueba
    'objConn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; pwd=mastercard;Initial Catalog=daniel;Data Source=VISITEC-PC"
    
    'Conexion con parametro Output y BD SENDA
    objConn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; pwd=mastercard;Initial Catalog=SENDAEFACT_LOCAL;Data Source=(LOCAL)"
    objConn.Open ' abrimos nuestra conexion
    '-------------------
    '--- Asignamos el Objeto comando para ejecutar store procedure
    '-------------------
    Set objCmd = New ADODB.Command
    objCmd.CommandType = adCmdStoredProc ' indicamos que es de tipo SP
    
    'objCmd.CommandText = "InsertarUsuario" aqui para BD de prueba daniel
    objCmd.CommandText = "SENDAEFACT_FIRMAR_XML_ARG"    ' BD SENDA aca debemos indicar el nombre de SP que se va a tratar
    Set objCmd.ActiveConnection = objConn 'asignamo la conexion a la BD al comando
     
    '-------------------
    '--- Creamos y asignamos los parametros
    '-------------------
    'En esta Seccion se debera indicar cada uno de los parametro segun el requerimiento de cada SP
    'DECLARE @RC int
    
    objCmd.Parameters.Append objCmd.CreateParameter("@RUC", adChar, adParamInput, 11, varRUC_EMISOR) 'RUC EMISOR
    objCmd.Parameters.Append objCmd.CreateParameter("@NRO_EFACT", adVarChar, adParamInput, 12, varNRO_EFACT) 'NUMERO DE COMPROBANTE DE PAGO EJEM: F00100000001 (SERIE =4 CARACTERES)(CORRELATIVO 8 CARACTERES)
    objCmd.Parameters.Append objCmd.CreateParameter("@TIDO", adChar, adParamInput, 2, varTIPODOCU) ' TIPO DE COMPROBANTE DE PAGO (01,03,07,08)
       
    '-------------------
    '--- Ejecutamos el Store Procedure
    '-------------------
    objCmd.Execute
    '-------------------
  
    
    
    'txtUserName.Text = Trim(objCmd.Parameters(70).Value)
    'txtClave.Text = objCmd.Parameters(71).Value
    
    'Para el ejemplo capturamos la salida o respuesta output que devuelve el SP y capturamos
    'su valor segun la matrix de variables y lo mostramos en unos Txtbox de Ejemplo pudiendo usarse para guardar en una BD
    'If IsNull(objCmd.Parameters(24).Value) Then
    '    varSTR_ERROR = "String Error Valor Nulo pero exitoso"
    'Else
    '    varSTR_ERROR = objCmd.Parameters(24).Value
    'End If

    'If IsNull(objCmd.Parameters(25).Value) Then
    '    varCOD_ERROR = -53
    'Else
    '    varCOD_ERROR = objCmd.Parameters(25).Value
    'End If
    
    'FrmPrincipal.txtUserName.Text = varSTR_ERROR
    'FrmPrincipal.txtClave.Text = varCOD_ERROR
    '-------------------
    
    '-------------------
    '----- Le asignamos nada al comando y conexion y se cierra
    objConn.Close
    Set objCmd = Nothing
    Set objConn = Nothing
    '-------------
    'MsgBox ("Documento SENDAEFACT_FIRMAR_XML_ARG"), vbOKOnly, "Firma XML_ARG"
End Sub

Sub SENDAEFACT_GET_DSVALUE()
    ' Variables Capturador de Error
    Dim varSTR_ERROR As String
    Dim varCOD_ERROR As Integer
    
    Dim objConn      As ADODB.Connection 'variable Objeto de conexion a la BD
    Dim objCmd       As ADODB.Command    'variable comando de ejecucion SP (Store Procedure)
    '-------------------
    '--- Abrir Coneccion con la BD
    '-------------------
    Set objConn = New ADODB.Connection
    objConn.Provider = "SQLOLEDB"
    'Conexion sencilla de prueba
    'objConn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; pwd=mastercard;Initial Catalog=daniel;Data Source=VISITEC-PC"
    
    'Conexion con parametro Output y BD SENDA
    objConn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; pwd=mastercard;Initial Catalog=SENDAEFACT_LOCAL;Data Source=(LOCAL)"
    objConn.Open ' abrimos nuestra conexion
    '-------------------
    '--- Asignamos el Objeto comando para ejecutar store procedure
    '-------------------
    Set objCmd = New ADODB.Command
    objCmd.CommandType = adCmdStoredProc ' indicamos que es de tipo SP
    
    'objCmd.CommandText = "InsertarUsuario" aqui para BD de prueba daniel
    objCmd.CommandText = "SENDAEFACT_GET_DSVALUE"    ' BD SENDA aca debemos indicar el nombre de SP que se va a tratar
    Set objCmd.ActiveConnection = objConn 'asignamo la conexion a la BD al comando
     
    '-------------------
    '--- Creamos y asignamos los parametros
    '-------------------
    'En esta Seccion se debera indicar cada uno de los parametro segun el requerimiento de cada SP
    'DECLARE @RC int
    
    objCmd.Parameters.Append objCmd.CreateParameter("@RUC", adChar, adParamInput, 11, varRUC_EMISOR) 'RUC EMISOR
    objCmd.Parameters.Append objCmd.CreateParameter("@NRO_EFACT", adVarChar, adParamInput, 12, varNRO_EFACT) 'NUMERO DE COMPROBANTE DE PAGO EJEM: F00100000001 (SERIE =4 CARACTERES)(CORRELATIVO 8 CARACTERES)
    objCmd.Parameters.Append objCmd.CreateParameter("@TIDO", adChar, adParamInput, 2, varTIPODOCU) ' TIPO DE COMPROBANTE DE PAGO (01,03,07,08)
    objCmd.Parameters.Append objCmd.CreateParameter("@DS_VALUE", adVarChar, adParamOutput, 50, InCadena) 'DESCRIPCION DEL ERROR
       
    '-------------------
    '--- Ejecutamos el Store Procedure
    '-------------------
    objCmd.Execute
    '-------------------
    
    varSTR_DS_VALUE = objCmd.Parameters(3)
    'txtUserName.Text = Trim(objCmd.Parameters(70).Value)
    'txtClave.Text = objCmd.Parameters(71).Value
    
    'Para el ejemplo capturamos la salida o respuesta output que devuelve el SP y capturamos
    'su valor segun la matrix de variables y lo mostramos en unos Txtbox de Ejemplo pudiendo usarse para guardar en una BD
    'If IsNull(objCmd.Parameters(24).Value) Then
    '    varSTR_ERROR = "String Error Valor Nulo pero exitoso"
    'Else
    '    varSTR_ERROR = objCmd.Parameters(24).Value
    'End If

    'If IsNull(objCmd.Parameters(25).Value) Then
    '    varCOD_ERROR = -53
    'Else
    '    varCOD_ERROR = objCmd.Parameters(25).Value
    'End If
    
    'FrmPrincipal.txtUserName.Text = varSTR_ERROR
    'FrmPrincipal.txtClave.Text = varCOD_ERROR
    '-------------------
    
    '-------------------
    '----- Le asignamos nada al comando y conexion y se cierra
    objConn.Close
    Set objCmd = Nothing
    Set objConn = Nothing
    '-------------
    'MsgBox ("Documento SENDAEFACT_GET_DSVALUE"), vbOKOnly, "EFACT_GET_DSVALUE"
End Sub

Sub SENDAEFACT_INSERT_DOCUMENTOBAJA()
    ' Variables Capturador de Error
    Dim varSTR_ERROR As String
    Dim varCOD_ERROR As Integer
    bvarFBAJA = Format(Now, "yyyy/MM/dd")
    bvarMOTIVO_BAJA = "BAJA"
    
    Dim objConn      As ADODB.Connection 'variable Objeto de conexion a la BD
    Dim objCmd       As ADODB.Command    'variable comando de ejecucion SP (Store Procedure)
    '-------------------
    '--- Abrir Coneccion con la BD
    '-------------------
    Set objConn = New ADODB.Connection
    objConn.Provider = "SQLOLEDB"
    'Conexion sencilla de prueba
    'objConn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; pwd=mastercard;Initial Catalog=daniel;Data Source=VISITEC-PC"
    
    'Conexion con parametro Output y BD SENDA
    objConn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; pwd=mastercard;Initial Catalog=SENDAEFACT_LOCAL;Data Source=(LOCAL)"
    objConn.Open ' abrimos nuestra conexion
    '-------------------
    '--- Asignamos el Objeto comando para ejecutar store procedure
    '-------------------
    Set objCmd = New ADODB.Command
    objCmd.CommandType = adCmdStoredProc ' indicamos que es de tipo SP
    
    'objCmd.CommandText = "InsertarUsuario" aqui para BD de prueba daniel
    objCmd.CommandText = "SENDAEFACT_INSERT_DOCUMENTOBAJA"    ' BD SENDA aca debemos indicar el nombre de SP que se va a tratar
    Set objCmd.ActiveConnection = objConn 'asignamo la conexion a la BD al comando
     
    '-------------------
    '--- Creamos y asignamos los parametros
    '-------------------
    'En esta Seccion se debera indicar cada uno de los parametro segun el requerimiento de cada SP
    'DECLARE @RC int
    
    objCmd.Parameters.Append objCmd.CreateParameter("@RUC_EMISOR", adChar, adParamInput, 11, varRUC_EMISOR) 'RUC EMISOR
    objCmd.Parameters.Append objCmd.CreateParameter("@TIPODOCU", adChar, adParamInput, 2, varTIPODOCU) ' TIPO DE COMPROBANTE DE PAGO (01,03,07,08)
    objCmd.Parameters.Append objCmd.CreateParameter("@NRO_EFACT", adVarChar, adParamInput, 12, varNRO_EFACT) 'NUMERO DE COMPROBANTE DE PAGO EJEM: F00100000001 (SERIE =4 CARACTERES)(CORRELATIVO 8 CARACTERES)
    objCmd.Parameters.Append objCmd.CreateParameter("@FBAJA", adDate, adParamInput, 10, bvarFBAJA) 'FECHA DE BAJA
    objCmd.Parameters.Append objCmd.CreateParameter("@MOTIVO_BAJA", adVarChar, adParamInput, 50, bvarMOTIVO_BAJA) ' DESCRIPCION DEL MOTIVO DE BAJA
    
    objCmd.Parameters.Append objCmd.CreateParameter("STR_ERROR", adVarChar, adParamOutput, 50, InCadena) 'DESCRIPCION DEL ERROR
    objCmd.Parameters.Append objCmd.CreateParameter("COD_ERROR", adInteger, adParamOutput, 3, InNumero) 'SI DEVUELTO -1 ES UN ERROR
    '-------------------
    '--- Ejecutamos el Store Procedure
    '-------------------
    objCmd.Execute
    '-------------------
    
    'txtUserName.Text = Trim(objCmd.Parameters(70).Value)
    'txtClave.Text = objCmd.Parameters(71).Value
    
    'Para el ejemplo capturamos la salida o respuesta output que devuelve el SP y capturamos
    'su valor segun la matrix de variables y lo mostramos en unos Txtbox de Ejemplo pudiendo usarse para guardar en una BD
    If IsNull(objCmd.Parameters(5).Value) Then
        varSTR_ERROR = "String Error Valor Nulo pero exitoso"
    Else
        varSTR_ERROR = objCmd.Parameters(5).Value
    End If

    If IsNull(objCmd.Parameters(6).Value) Then
        varCOD_ERROR = -53
    Else
        varCOD_ERROR = objCmd.Parameters(6).Value
    End If
    
    'FrmPrincipal.txtusername.Text = varSTR_ERROR
    'FrmPrincipal.txtClave.Text = varCOD_ERROR
    '-------------------
    
    '-------------------
    '----- Le asignamos nada al comando y conexion y se cierra
    objConn.Close
    Set objCmd = Nothing
    Set objConn = Nothing
    '-------------
   ' MsgBox ("Documento SENDAEFACT_INSERT_DOCUMENTOBAJA"), vbOKOnly, "SENDAEFACT_INSERT_DOCUMENTOBAJA"
End Sub

Sub SENDAEFACT_OBTENER_DS_VALUE()
    ' Variables Capturador de Error
    Dim varSTR_ERROR As String
    Dim varCOD_ERROR As Integer
    
    Dim objConn      As ADODB.Connection 'variable Objeto de conexion a la BD
    Dim objCmd       As ADODB.Command    'variable comando de ejecucion SP (Store Procedure)
    '-------------------
    '--- Abrir Coneccion con la BD
    '-------------------
    Set objConn = New ADODB.Connection
    objConn.Provider = "SQLOLEDB"
    'Conexion sencilla de prueba
    'objConn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; pwd=mastercard;Initial Catalog=daniel;Data Source=VISITEC-PC"
    
    'Conexion con parametro Output y BD SENDA
    objConn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; pwd=mastercard;Initial Catalog=SENDAEFACT_LOCAL;Data Source=(LOCAL)"
    objConn.Open ' abrimos nuestra conexion
    '-------------------
    '--- Asignamos el Objeto comando para ejecutar store procedure
    '-------------------
    Set objCmd = New ADODB.Command
    objCmd.CommandType = adCmdStoredProc ' indicamos que es de tipo SP
    
    'objCmd.CommandText = "InsertarUsuario" aqui para BD de prueba daniel
    objCmd.CommandText = "OBTENER_DS_VALUE"    ' BD SENDA aca debemos indicar el nombre de SP que se va a tratar
    Set objCmd.ActiveConnection = objConn 'asignamo la conexion a la BD al comando
     
    '-------------------
    '--- Creamos y asignamos los parametros
    '-------------------
    'En esta Seccion se debera indicar cada uno de los parametro segun el requerimiento de cada SP
    'DECLARE @RC int
    
    objCmd.Parameters.Append objCmd.CreateParameter("@RUC", adChar, adParamInput, 11, varRUC_EMISOR) 'RUC EMISOR
    objCmd.Parameters.Append objCmd.CreateParameter("@NRO_EFACT", adVarChar, adParamInput, 12, varNRO_EFACT) 'NUMERO DE COMPROBANTE DE PAGO EJEM: F00100000001 (SERIE =4 CARACTERES)(CORRELATIVO 8 CARACTERES)
    objCmd.Parameters.Append objCmd.CreateParameter("@TIPO_DOC", adChar, adParamInput, 2, varTIPODOCU) ' TIPO DE COMPROBANTE DE PAGO (01,03,07,08)
    objCmd.Parameters.Append objCmd.CreateParameter("@DS_VALUE", adVarChar, adParamOutput, 50, InCadena) 'DESCRIPCION DEL ERROR
       
    '-------------------
    '--- Ejecutamos el Store Procedure
    '-------------------
    objCmd.Execute
    '-------------------
    ' varSTR_ERROR = objCmd.Parameters(3)
    varSTR_DS_VALUE = objCmd.Parameters(3)
    
    'txtUserName.Text = Trim(objCmd.Parameters(70).Value)
    'txtClave.Text = objCmd.Parameters(71).Value
    
    'Para el ejemplo capturamos la salida o respuesta output que devuelve el SP y capturamos
    'su valor segun la matrix de variables y lo mostramos en unos Txtbox de Ejemplo pudiendo usarse para guardar en una BD
    'If IsNull(objCmd.Parameters(24).Value) Then
    '    varSTR_ERROR = "String Error Valor Nulo pero exitoso"
    'Else
    '    varSTR_ERROR = objCmd.Parameters(24).Value
    'End If

    'If IsNull(objCmd.Parameters(25).Value) Then
    '    varCOD_ERROR = -53
    'Else
    '    varCOD_ERROR = objCmd.Parameters(25).Value
    'End If
    
    'FrmPrincipal.txtUserName.Text = varSTR_ERROR
    'FrmPrincipal.txtClave.Text = varCOD_ERROR
    '-------------------
    
    '-------------------
    '----- Le asignamos nada al comando y conexion y se cierra
    objConn.Close
    Set objCmd = Nothing
    Set objConn = Nothing
    '-------------
    'MsgBox ("Documento SENDAEFACT_FIRMAR_XML_ARG"), vbOKOnly, "Firma XML_ARG"
End Sub


Sub SENDAEFACT_REPRINT()
    ' Variables Capturador de Error
    Dim varSTR_ERROR As String
    Dim varCOD_ERROR As Integer
    
    Dim objConn      As ADODB.Connection 'variable Objeto de conexion a la BD
    Dim objCmd       As ADODB.Command    'variable comando de ejecucion SP (Store Procedure)
    '-------------------
    '--- Abrir Coneccion con la BD
    '-------------------
    Set objConn = New ADODB.Connection
    objConn.Provider = "SQLOLEDB"
    'Conexion sencilla de prueba
    'objConn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; pwd=mastercard;Initial Catalog=daniel;Data Source=VISITEC-PC"
    
    'Conexion con parametro Output y BD SENDA
    objConn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; pwd=mastercard;Initial Catalog=SENDAEFACT_LOCAL;Data Source=(LOCAL)"
    objConn.Open ' abrimos nuestra conexion
    '-------------------
    '--- Asignamos el Objeto comando para ejecutar store procedure
    '-------------------
    Set objCmd = New ADODB.Command
    objCmd.CommandType = adCmdStoredProc ' indicamos que es de tipo SP
    
    'objCmd.CommandText = "InsertarUsuario" aqui para BD de prueba daniel
    objCmd.CommandText = "SENDAEFACT_REPRINT"    ' BD SENDA aca debemos indicar el nombre de SP que se va a tratar
        
    Set objCmd.ActiveConnection = objConn 'asignamo la conexion a la BD al comando
     
    '-------------------
    '--- Creamos y asignamos los parametros
    '-------------------
    'En esta Seccion se debera indicar cada uno de los parametro segun el requerimiento de cada SP
    'DECLARE @RC int
    
    'objConn.Open "SELECT * FROM vendedorcomision where  producto='" & xproducto & "' and codigo='" & xvendedor & "' ", cn, adOpenDynamic, adLockOptimistic
    
    objConn.Execute ("update DOCUMENTOVENTA set bprint=0 where ruc_emisor='" & RePrintRUC_EMISOR & "' and tipodocu='" & RePrintvarTIPODOCU & "' and nro_efact='" & RePrintvarNRO_EFACT & "' ")
    
    
    'objCmd.Parameters.Append objCmd.CreateParameter("@RUC", adChar, adParamInput, 11, RePrintRUC_EMISOR) 'RUC EMISOR
    'objCmd.Parameters.Append objCmd.CreateParameter("@NRO_EFACT", adVarChar, adParamInput, 12, RePrintvarNRO_EFACT) 'NUMERO DE COMPROBANTE DE PAGO EJEM: F00100000001 (SERIE =4 CARACTERES)(CORRELATIVO 8 CARACTERES)
    'objCmd.Parameters.Append objCmd.CreateParameter("@TIPO_DOC", adChar, adParamInput, 2, RePrintvarTIPODOCU) ' TIPO DE COMPROBANTE DE PAGO (01,03,07,08)
    
    'MsgBox ("Datos"    RePrintRUC_EMISOR    RePrintvarNRO_EFACT   RePrintvarTIPODOCU), vbOKOnly, "Copia Documento"
    MsgBox ("Datos Documento (" & RePrintRUC_EMISOR & ") (" & RePrintvarNRO_EFACT & ") (" & RePrintvarTIPODOCU & ") "), vbOKOnly, "Documento"

    '-------------------
    '--- Ejecutamos el Store Procedure
    '-------------------
    'objCmd.Execute
    '-------------------
    
    
    'Para el ejemplo capturamos la salida o respuesta output que devuelve el SP y capturamos
    'su valor segun la matrix de variables y lo mostramos en unos Txtbox de Ejemplo pudiendo usarse para guardar en una BD
    'If IsNull(objCmd.Parameters(24).Value) Then
    '    varSTR_ERROR = "String Error Valor Nulo pero exitoso"
    'Else
    '    varSTR_ERROR = objCmd.Parameters(24).Value
    'End If

    'If IsNull(objCmd.Parameters(25).Value) Then
    '    varCOD_ERROR = -53
    'Else
    '    varCOD_ERROR = objCmd.Parameters(25).Value
    'End If
    
    'FrmPrincipal.txtUserName.Text = varSTR_ERROR523865274
    'FrmPrincipal.txtClave.Text = varCOD_ERROR
    '-------------------
    
    '-------------------
    '----- Le asignamos nada al comando y conexion y se cierra
    objConn.Close
    Set objCmd = Nothing
    Set objConn = Nothing
    '-------------
    'MsgBox ("Documento SENDAEFACT_FIRMAR_XML_ARG"), vbOKOnly, "Firma XML_ARG"
End Sub




