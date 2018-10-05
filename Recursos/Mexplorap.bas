Attribute VB_Name = "Mexplorap"

'inicio 01/07/2017 pll
'para la grid consulta
Const N_cols_dflag = 0

Const N_cols_yausado = 1

Const N_cols_estado = 2

Const N_cols_local = 3

Const N_cols_tipo = 4

Const N_cols_serie = 5

Const N_cols_numero = 6

Const N_cols_fecha = 7

Const N_cols_fechae = 8

Const N_cols_hora = 9

Const N_cols_tipoclie = 10

Const N_cols_codigo = 11

Const N_cols_nombre = 12

Const N_cols_moneda = 13

Const N_cols_total = 14

Const N_cols_acuenta = 15

Const N_cols_adetotal = 16

Const N_cols_bodega = 17

Const N_cols_bodegaf = 18

Const N_cols_localf = 19

Const N_cols_nro_items = 20

Const N_cols_usuario = 21

Const N_cols_placa = 22

Const N_cols_caja = 23

Const N_cols_turno = 24

Const N_cols_vendedor = 25

Const N_cols_observa = 26

Const N_cols_acu = 27

Const N_cols_servicio = 28

Const N_cols_retipo1 = 29

Const N_cols_renumero3 = 30

Const N_cols_renumero1 = 31

Const N_cols_renumero2 = 32

Const N_cols_neto = 33

Const N_cols_descuento = 34

Const N_cols_subtotal = 35

Const N_cols_impuesto = 36

Const N_cols_tipoimp = 37

Const N_cols_acu1 = 38

Const N_cols_tipo1 = 39

Const N_cols_serie1 = 40

Const N_cols_numero1 = 41

Const N_cols_descrTipo = 42

Const N_cols_fpago = 43

Type struc_deta_cotizacion

    dflag                         As String
    yausado                       As String
    estado                        As String
    local                         As String
    tipo                          As String
    serie                         As String
    Numero                        As String
    fecha                         As String
    fechae                        As String
    tipoclie                      As String
    codigo                        As String
    nombre                        As String
    moneda                        As String
    total                         As Double
    acuenta                       As Double
    adetotal                      As Double
    bodega                        As String
    bodegaf                       As String
    localf                        As String
    nro_items                     As String
    usuario                       As String
    placa                         As String
    turno                         As String
    vendedor                      As String
    observa                       As String
    servicio                      As String
    local1                        As String
    tipo1                         As String
    serie1                        As String
    numero1                       As String
    retipo1                       As String
    renumero3                     As String
    renumero1                     As String
    renumero2                     As String
    neto                          As String
    descuento                     As Currency
    subtotal                      As Currency
    impuesto                      As Currency
    tipoimp                       As String
    acu1                          As String
    acu                           As String
    hora                          As String
    servicioco                    As String
    percepcion                    As String
    descriTipo                    As String
    fpago                         As String
    esunat                        As String
    caja                          As String

End Type

Type struc_nueva_cotizacion

    nombre                       As String

End Type

Global my_struc_deta_cotizacion()  As struc_deta_cotizacion

Global my_struc_nueva_cotizacion() As struc_nueva_cotizacion

'inicio 29/08/2017 pll
Global my_numero                   As String

Global my_local                    As String

Global my_tipo                     As String

Global my_serie                    As String

Public my_nombre                   As String

Public my_codigo                   As String

Public my_acu                      As String

Public my_acu1                     As String

Public my_descrTipo                As String

Public my_fecha                    As String

'esto es para la modificacion de cotizacion pll
Public my_estado                   As String

Public my_tipo1                    As String

Public my_serie1                   As String

Public my_numero1                  As String

Public my_renumero1                As String

Public my_renumero2                As String

Public my_renumero3                As String

Public my_retipo1                  As String

Public my_total                    As Double

Public local1                      As String

Public acu                         As String

Public my_yausado                  As String

Public my_turno                    As String

Public my_caja                     As String

Public my_vendedor                 As String

Public my_moneda                   As String

'esto es para generar
Public my_bodegaf                  As String

Public my_bodega                   As String

Public my_localf                   As String

Public my_tipoclie                 As String

Public my_fpago                    As String * 1

'esto es para consulta-condiciones
Public my_servicio                 As String
'Public my_placa                       As String
'para la forma de pago 02/07/2017 pll

' Testing Proyecto Facturacion Electronica
Public my_estadosunat              As String
' Testing Proyecto Facturacion Electronica

Const N_colsm_local = 0

Const N_colsm_tipo = 1

Const N_colsm_serie = 2

Const N_colsm_numero = 3

Const N_colsm_fpago = 4

Const N_colsm_descripcion = 5

Const N_colsm_moneda = 6

Const N_colsm_total = 7

Const N_colsm_entrega = 8

Const N_colsm_saldos = 9

Type struc_fpago

    local                         As String
    tipo                          As String
    serie                         As String
    Numero                        As String
    fpago                         As String
    descripcion                   As String
    moneda                        As String
    total                         As Double
    entrega                       As String
    saldos                        As Double

End Type

Global my_struc_fpago() As struc_fpago

'inicio 03/08/2017 pll

Type struc_solo_documentos

    codigo                         As String
    nombre                         As String
    local                          As String
    estado                         As String
    tipo                           As String
    serie                          As String
    Numero                         As String
    fecha                          As String
    total                          As Double

End Type

Global my_struc_solo_documentos() As struc_solo_documentos

'fin 03/08/2017 pll
Public Function ini_grid_dcotizacion(my_grid As MSFlexGrid)

    'utilizo per eliminare_fattura<--voce
    my_grid.Clear
    'inizializzazione Grid
    my_grid.rows = 2
    my_grid.FixedRows = 1
    my_grid.Cols = 45 'aqui se aumenta
    my_grid.FixedCols = 0

    my_grid.Row = 0
    my_grid.Col = N_cols_dflag 'dflag
    my_grid.Text = "X"
    my_grid.ColWidth(my_grid.Col) = 50
    my_grid.FixedAlignment(my_grid.Col) = 0
    my_grid.ColAlignment(my_grid.Col) = 0

    my_grid.Row = 0
    my_grid.Col = N_cols_yausado 'yausado
    my_grid.Text = "A"
    my_grid.ColWidth(my_grid.Col) = 250
    my_grid.FixedAlignment(my_grid.Col) = 0
    my_grid.ColAlignment(my_grid.Col) = 0

    my_grid.Row = 0
    my_grid.Col = N_cols_estado 'estado 1
    my_grid.Text = "Estado"
    my_grid.ColWidth(my_grid.Col) = 500
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_local 'estado 2
    my_grid.Text = "Local"
    my_grid.ColWidth(my_grid.Col) = 500
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_tipo 'tipo 3
    my_grid.Text = "Tipo"
    my_grid.ColWidth(my_grid.Col) = 500
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_serie 'serie 4
    my_grid.Text = "Serie"
    my_grid.ColWidth(my_grid.Col) = 900
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_numero 'numero 5
    my_grid.Text = "Numero"
    my_grid.ColWidth(my_grid.Col) = 1200
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_fecha 'fecha 6
    my_grid.Text = "Fecha"
    my_grid.ColWidth(my_grid.Col) = 1800
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_fechae 'fechae 7
    my_grid.Text = "Fechae"
    my_grid.ColWidth(my_grid.Col) = 1800
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_hora 'hora 8
    my_grid.Text = "Hora"
    my_grid.ColWidth(my_grid.Col) = 1200
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_tipoclie 'tipoclie 9
    my_grid.Text = "Tipo cliente"
    my_grid.ColWidth(my_grid.Col) = 1200
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_codigo 'codigo 10
    my_grid.Text = "Codigo"
    my_grid.ColWidth(my_grid.Col) = 1600
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_nombre 'nombre 11
    my_grid.Text = "Nombre"
    my_grid.ColWidth(my_grid.Col) = 2700
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_moneda 'moneda 12
    my_grid.Text = "Moneda"
    my_grid.ColWidth(my_grid.Col) = 800
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_total 'total 13
    my_grid.Text = "Total"
    my_grid.ColWidth(my_grid.Col) = 1200
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_acuenta 'acuenta 14
    my_grid.Text = "Acuenta"
    my_grid.ColWidth(my_grid.Col) = 1200
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_adetotal 'adetotal 15
    my_grid.Text = "Adetotal"
    my_grid.ColWidth(my_grid.Col) = 1200
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_bodega 'bodega 16
    my_grid.Text = "Bodega"
    my_grid.ColWidth(my_grid.Col) = 800
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_bodegaf 'bodegaf 17
    my_grid.Text = "Bodegaf"
    my_grid.ColWidth(my_grid.Col) = 500
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_localf 'localf 18
    my_grid.Text = "Localf"
    my_grid.ColWidth(my_grid.Col) = 500
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_nro_items 'nro_items 19
    my_grid.Text = "Nro Items"
    my_grid.ColWidth(my_grid.Col) = 800
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_usuario 'usuario 20
    my_grid.Text = "Usuario"
    my_grid.ColWidth(my_grid.Col) = 500
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_placa 'placa 21
    my_grid.Text = "Placa"
    my_grid.ColWidth(my_grid.Col) = 500
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_caja 'caja 22
    my_grid.Text = "Caja"
    my_grid.ColWidth(my_grid.Col) = 500
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_turno 'turno 23
    my_grid.Text = "Turno"
    my_grid.ColWidth(my_grid.Col) = 500
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_vendedor 'vendedor 24
    my_grid.Text = "Vendedor"
    my_grid.ColWidth(my_grid.Col) = 1200
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_observa 'observa 25
    my_grid.Text = "Observa"
    my_grid.ColWidth(my_grid.Col) = 800
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_acu 'acu 27
    my_grid.Text = "Acu"
    my_grid.ColWidth(my_grid.Col) = 800
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_servicio 'servicio 28
    my_grid.Text = "Servicio"
    my_grid.ColWidth(my_grid.Col) = 800
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_retipo1 'retipo1 33
    my_grid.Text = "Retipo 1"
    my_grid.ColWidth(my_grid.Col) = 800
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_renumero3 'renumero3 34
    my_grid.Text = "Renumero 3"
    my_grid.ColWidth(my_grid.Col) = 800
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_renumero1 'renumero1 35
    my_grid.Text = "Renumero 1"
    my_grid.ColWidth(my_grid.Col) = 800
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_renumero2 'renumero2 36
    my_grid.Text = "Renumero 2"
    my_grid.ColWidth(my_grid.Col) = 800
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_neto 'neto 37
    my_grid.Text = "Neto"
    my_grid.ColWidth(my_grid.Col) = 800
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_descuento 'neto descuento 38
    my_grid.Text = "Descuento"
    my_grid.ColWidth(my_grid.Col) = 800
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_subtotal 'subtotal 39
    my_grid.Text = "Subtotal"
    my_grid.ColWidth(my_grid.Col) = 800
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_impuesto 'impuesto 40
    my_grid.Text = "Impuesto"
    my_grid.ColWidth(my_grid.Col) = 1200
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_tipoimp 'tipoimp 41
    my_grid.Text = "Tipo Imp "
    my_grid.ColWidth(my_grid.Col) = 1200
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_acu1 'acu1 42
    my_grid.Text = "Acu 1"
    my_grid.ColWidth(my_grid.Col) = 800
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    'aqui no visualiza los datos
    my_grid.Row = 0
    my_grid.Col = N_cols_tipo1 'Tipo1 43
    my_grid.Text = "Tipo1"
    my_grid.ColWidth(my_grid.Col) = 800
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_serie1 'serie1 44
    my_grid.Text = "serie1"
    my_grid.ColWidth(my_grid.Col) = 800
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_numero1 'numero1 45
    my_grid.Text = "numero1"
    my_grid.ColWidth(my_grid.Col) = 800
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_descrTipo 'numero1 46
    my_grid.Text = "descriTipo" 'descripcion del tipo documento
    my_grid.ColWidth(my_grid.Col) = 0
    my_grid.FixedAlignment(my_grid.Col) = 0
    my_grid.ColAlignment(my_grid.Col) = 0

    my_grid.Row = 0
    my_grid.Col = N_cols_fpago 'numero1 47
    my_grid.Text = "fpago" 'forma de pago
    my_grid.ColWidth(my_grid.Col) = 0
    my_grid.FixedAlignment(my_grid.Col) = 0
    my_grid.ColAlignment(my_grid.Col) = 0

    Exit Function

End Function

Public Function carica_deta_cotizacion(my_grid As MSFlexGrid, _
                                       my_struc_deta_cotizacion() As struc_deta_cotizacion, _
                                       k As Integer, _
                                       my_total As Double, _
                                       my_neto As Double, _
                                       my_subtotal As Double, _
                                       my_impuesto As Double, _
                                       my_lvendido As Integer, _
                                       my_anulado As Integer, _
                                       my_tota_anulado As Double, _
                                       my_servicioco As Integer, _
                                       my_percepcion As Integer, _
                                       my_total_agravados As Double, _
                                       my_nume_agravados As Integer)

    For I = 0 To k - 1
        my_grid.AddItem ""
        my_grid.Row = my_grid.rows - 1

        my_grid.Col = N_cols_dflag
        my_grid.Text = my_struc_deta_cotizacion(I).dflag

        my_grid.Col = N_cols_yausado
        my_grid.Text = my_struc_deta_cotizacion(I).yausado

        my_grid.Col = N_cols_estado
        my_grid.Text = my_struc_deta_cotizacion(I).estado

        If my_struc_deta_cotizacion(I).estado = "2" Then
            my_lvendido = my_lvendido + 1

        End If

        If my_struc_deta_cotizacion(I).estado = "1" Then
            my_anulado = my_anulado + 1
            my_tota_anulado = my_tota_anulado + my_struc_deta_cotizacion(I).total

        End If

        If my_struc_deta_cotizacion(I).estado = "0" Then
            my_nume_agravados = my_nume_agravados + 1
            my_total_agravados = my_total_agravados + my_struc_deta_cotizacion(I).estado

        End If

        my_grid.Col = N_cols_local
        my_grid.Text = my_struc_deta_cotizacion(I).local

        my_grid.Col = N_cols_tipo
        my_grid.Text = my_struc_deta_cotizacion(I).tipo

        my_grid.Col = N_cols_serie
        my_grid.Text = my_struc_deta_cotizacion(I).serie

        my_grid.Col = N_cols_numero
        my_grid.Text = my_struc_deta_cotizacion(I).Numero

        my_grid.Col = N_cols_fecha
        my_grid.Text = my_struc_deta_cotizacion(I).fecha

        my_grid.Col = N_cols_fechae
        my_grid.Text = my_struc_deta_cotizacion(I).fechae

        my_grid.Col = N_cols_hora
        my_grid.Text = my_struc_deta_cotizacion(I).hora

        my_grid.Col = N_cols_tipoclie
        my_grid.Text = my_struc_deta_cotizacion(I).tipoclie

        my_grid.Col = N_cols_codigo
        my_grid.Text = my_struc_deta_cotizacion(I).codigo

        my_grid.Col = N_cols_nombre
        my_grid.Text = my_struc_deta_cotizacion(I).nombre

        my_grid.Col = N_cols_moneda
        my_grid.Text = my_struc_deta_cotizacion(I).moneda

        my_grid.Col = N_cols_total
        my_grid.Text = my_struc_deta_cotizacion(I).total
        my_total = my_total + my_struc_deta_cotizacion(I).total

        my_grid.Col = N_cols_acuenta
        my_grid.Text = my_struc_deta_cotizacion(I).acuenta

        my_grid.Col = N_cols_adetotal
        my_grid.Text = my_struc_deta_cotizacion(I).adetotal

        my_grid.Col = N_cols_bodega
        my_grid.Text = my_struc_deta_cotizacion(I).bodega

        my_grid.Col = N_cols_bodegaf
        my_grid.Text = my_struc_deta_cotizacion(I).bodegaf

        my_grid.Col = N_cols_localf
        my_grid.Text = my_struc_deta_cotizacion(I).localf

        my_grid.Col = N_cols_nro_items
        my_grid.Text = my_struc_deta_cotizacion(I).nro_items

        my_grid.Col = N_cols_usuario
        my_grid.Text = my_struc_deta_cotizacion(I).usuario

        my_grid.Col = N_cols_placa
        my_grid.Text = my_struc_deta_cotizacion(I).placa

        my_grid.Col = N_cols_caja
        my_grid.Text = my_struc_deta_cotizacion(I).caja

        my_grid.Col = N_cols_turno
        my_grid.Text = my_struc_deta_cotizacion(I).turno

        my_grid.Col = N_cols_vendedor
        my_grid.Text = my_struc_deta_cotizacion(I).vendedor

        my_grid.Col = N_cols_observa
        my_grid.Text = my_struc_deta_cotizacion(I).observa

        my_grid.Col = N_cols_acu
        my_grid.Text = my_struc_deta_cotizacion(I).acu

        my_grid.Col = N_cols_servicio
        my_grid.Text = my_struc_deta_cotizacion(I).servicio

        my_grid.Col = N_cols_retipo1
        my_grid.Text = my_struc_deta_cotizacion(I).retipo1

        my_grid.Col = N_cols_renumero3
        my_grid.Text = my_struc_deta_cotizacion(I).renumero3

        my_grid.Col = N_cols_renumero1
        my_grid.Text = my_struc_deta_cotizacion(I).renumero1

        my_grid.Col = N_cols_renumero1
        my_grid.Text = my_struc_deta_cotizacion(I).renumero2

        my_grid.Col = N_cols_neto
        my_grid.Text = my_struc_deta_cotizacion(I).neto
        my_neto = my_neto + my_struc_deta_cotizacion(I).neto

        my_grid.Col = N_cols_descuento
        my_grid.Text = my_struc_deta_cotizacion(I).descuento

        my_grid.Col = N_cols_subtotal
        my_grid.Text = my_struc_deta_cotizacion(I).subtotal
        my_subtotal = my_subtotal + my_struc_deta_cotizacion(I).subtotal

        my_grid.Col = N_cols_impuesto
        my_grid.Text = my_struc_deta_cotizacion(I).impuesto
        my_impuesto = my_impuesto + my_struc_deta_cotizacion(I).impuesto

        my_grid.Col = N_cols_tipoimp
        my_grid.Text = my_struc_deta_cotizacion(I).tipoimp

        my_grid.Col = N_cols_acu1
        my_grid.Text = my_struc_deta_cotizacion(I).acu1

        'datos que no son visualizados pero considerados pll
        If my_struc_deta_cotizacion(I).servicioco = "" Then
            my_servicioco = my_servicioco + 0
        Else
            my_servicioco = my_servicioco + my_struc_deta_cotizacion(I).servicioco

        End If

        If my_struc_deta_cotizacion(I).percepcion = "" Then
            my_percepcion = my_percepcion + 0
        Else
            my_percepcion = my_percepcion + my_struc_deta_cotizacion(I).percepcion

        End If

        'aqui no visualiza los datos pero son necesarios para elaborar la modificacion pll
        my_grid.Col = N_cols_tipo1
        my_grid.Text = my_struc_deta_cotizacion(I).tipo1

        my_grid.Col = N_cols_serie1
        my_grid.Text = my_struc_deta_cotizacion(I).serie1

        my_grid.Col = N_cols_numero1
        my_grid.Text = my_struc_deta_cotizacion(I).numero1

        my_grid.Col = N_cols_renumero1
        my_grid.Text = my_struc_deta_cotizacion(I).renumero1

        my_grid.Col = N_cols_renumero2
        my_grid.Text = my_struc_deta_cotizacion(I).renumero2

        my_grid.Col = N_cols_renumero3
        my_grid.Text = my_struc_deta_cotizacion(I).renumero3

        'inicio 16/08/2017 pll para la descripcion tipo
        my_grid.Col = N_cols_descrTipo
        my_grid.Text = my_struc_deta_cotizacion(I).descriTipo
        'fin 16/08/2017 pll
        'inicio 07/09/2017 pll para la forma de pago
        my_grid.Col = N_cols_fpago
        my_grid.Text = my_struc_deta_cotizacion(I).fpago
        'fin 07/09/2017 pll para la forma de pago

    Next I

    my_grid.Row = 1
    my_grid.Col = 0

    If my_grid.Text = "" Then
        my_grid.RemoveItem 1

    End If

    Exit Function

End Function

Sub sql_cabeza(my_local As String, _
               my_tipo As String, _
               my_caja As String, _
               my_vendedor As String, _
               my_cajero As String, _
               my_bodega As String, _
               my_bodegaf As String, _
               my_servicio As String, _
               my_serie As String, _
               my_numero As String, _
               my_fechai As String, _
               my_fechaf As String, _
               acu As String, _
               my_combo2 As String, _
               my_ordenado As String, _
               my_struc_deta_cotizacion() As struc_deta_cotizacion, _
               salida As Boolean, _
               k As Integer, _
               my_moneda As String, _
               my_codigo As String, _
               my_nombre As String, _
               my_estado As String, _
               my_placa As String, _
               my_saldo_inicial As Boolean)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd921_err

    If Len(my_fechai) <> 10 Then Exit Sub
    If Len(my_fechaf) <> 10 Then Exit Sub
    If Not IsDate(my_fechai) Then Exit Sub
    If Not IsDate(my_fechaf) Then Exit Sub
    'MsgBox cgusuario
    ReDim my_struc_deta_cotizacion(0)

    mysql = "select f.*,t.DESCRIPCIO " & Chr$(10)
    mysql = mysql & "from " & cgusuario & " f," & Chr$(10)
    mysql = mysql & "tipo t" & Chr$(10)
    mysql = mysql & "where " & Chr$(10)

    If ve = "V" Then
        mysql = mysql & "f.fechae>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and f.fechae<='" & Format(my_fechaf, "YYYYMMDD") & "' " & Chr$(10)
    Else
        mysql = mysql & "f.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and f.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "' " & Chr$(10)

    End If

    If Trim(my_local) <> "%" Then
        mysql = mysql & " and f.local like '" & extra_loquesea(my_local) & "'" & Chr$(10)

    End If

    If Trim(my_tipo) <> "%" Then
        mysql = mysql & " and f.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)

    End If

    If Trim(my_caja) <> "%" Then
        mysql = mysql & " and f.caja like '" & extra_loquesea(my_caja) & "'" & Chr$(10)

    End If

    If Trim(my_turno) <> "%" Then
        mysql = mysql & " and f.turno like '" & extra_loquesea(my_turno) & "'" & Chr$(10)

    End If

    If my_serie <> "%" Then
        mysql = mysql & " and f.serie like '" & extra_loquesea(my_serie) & "'" & Chr$(10)

    End If

    If my_numero <> "%" Then
        mysql = mysql & " and f.numero like '" & extra_loquesea(my_numero) & "'" & Chr$(10)

    End If

    If my_codigo <> "%" Then
        mysql = mysql & " and f.codigo like '" & my_codigo & "'" & Chr$(10)

    End If

    If my_nombre <> "%" Then
        mysql = mysql & " and f.nombre like '" & my_nombre & "'" & Chr$(10)

    End If

    If my_moneda <> "%" Then
        mysql = mysql & " and f.moneda like '" & my_moneda & "'" & Chr$(10)

    End If

    If my_estado <> "%" Then
        mysql = mysql & " and f.estado like '" & my_estado & "'" & Chr$(10)

    End If

    If my_placa <> "%" Then
        mysql = mysql & " and f.placa='" & my_placa & "'" & Chr$(10)

    End If

    If Trim(my_vendedor) <> "%" Then
        mysql = mysql & " and f.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)

    End If

    If Trim(my_cajero) <> "%" Then
        mysql = mysql & " and f.usuario like '" & extra_loquesea(my_cajero) & "'" & Chr$(10)

    End If

    If Trim(my_bodega) <> "%" Then
        mysql = mysql & " and f.bodega like '" & extra_loquesea(my_bodega) & "'" & Chr$(10)

    End If

    If Trim(my_bodegaf) <> "%" Then
        mysql = mysql & " and f.bodegaf like '" & extra_loquesea(my_bodegaf) & "'" & Chr$(10)

    End If

    If Trim(my_servicio) <> "%" Then
        mysql = mysql & " or  f.servicio='" & my_servicio & "'" & Chr$(10)

    End If

    If my_saldo_inicial = True Then
        mysql = mysql & " and f.nop='S' " & Chr$(10)

    End If

    If acu <> "C" And acu <> "V" And acu <> "T" Then
        mysql = mysql & " and f.acu='" & acu & "'" & Chr$(10)
        'inicio 16/08/2017 pll para la descripcion tipo
        mysql = mysql & " and f.tipo = t.tipo" & Chr$(10)

    End If

    If my_combo2 <> "%" Then
        If my_combo2 = "Atendido" Then
            mysql = mysql & " and  f.yausado='1'" & Chr$(10)

        End If

        If my_combo2 = "Pendiente" Then
            mysql = mysql & " and  f.yausado='0'" & Chr$(10)

        End If

    End If

    If acu = "V" Then
        '19/06/2017 kenyo NOTA DE CREDITO
        'mysql = mysql & " and (f.acu='A' OR f.acu='B' OR f.acu='C' OR f.acu='D' OR f.acu='G' )" & Chr$(10)
        mysql = mysql & " and (acu='A' OR acu='B' OR acu='C' OR acu='D' OR acu='G' or acu='E' )" & Chr$(10)
        '19/06/2017 kenyo NOTA DE CREDITO

        mysql = mysql & " and f.tipo = t.tipo" & Chr$(10)

        If explorap.Check1.Value = 1 Then
            mysql = mysql & " and f.tipo<>'5'" & Chr$(10)

        End If

    End If

    If acu = "C" Then
        mysql = mysql & " and (f.acu='J' OR f.acu='K' OR f.acu='L' OR f.acu='M' OR f.acu='P' )" & Chr$(10)
        mysql = mysql & " and f.tipo = t.tipo" & Chr$(10)

    End If

    If acu = "T" Then
        'mysql = mysql & " and (f.acu='T' OR f.acu='S' )" & Chr$(10)
        mysql = mysql & " and (f.acu='T')" & Chr$(10)
        mysql = mysql & " and f.tipo = t.tipo" & Chr$(10)

    End If

    If acu = "GE" Then
        'mysql = mysql & " and (f.acu='T' OR f.acu='S' )" & Chr$(10)
        mysql = mysql & " and (f.acu='GE')" & Chr$(10)
        mysql = mysql & " and f.tipo = t.tipo" & Chr$(10)

    End If

    If acu = "S" Then
        mysql = mysql & " and (f.acu='S' )" & Chr$(10)
        mysql = mysql & " and f.tipo = t.tipo" & Chr$(10)

    End If

    If importacion = "IMPORTACION" Then
        mysql = mysql & " and f.tipoimp='I'" & Chr$(10)

    End If

    If importacion = "GASTOS" Then
        mysql = mysql & " and f.tipoimp='G'" & Chr$(10)

    End If

    If importacion = "COMERCIAL" Then
        mysql = mysql & " and (f.tipoimp='C' or f.tipoimp is null) " & Chr$(10)

    End If

    If my_ordenado <> "%" Then
        mysql = mysql & "order by " & my_ordenado & Chr$(10)
    Else
        my_ordenado = "f.fecha desc,f.hora desc"
        mysql = mysql & "order by " & my_ordenado & Chr$(10)

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

            'my_moneda = mytablex.Fields("moneda") '04/09/2017 pll
            'para cargar la msgrid
            If k > 0 Then
                ReDim Preserve my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion) + 1)

            End If

            If mytablex.Fields("dflag") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).dflag = mytablex.Fields("dflag")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).dflag = 0

            End If

            If mytablex.Fields("yausado") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).yausado = mytablex.Fields("yausado")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).yausado = 0

            End If

            If mytablex.Fields("estado") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).estado = mytablex.Fields("estado")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).estado = ""

            End If

            If mytablex.Fields("local") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).local = mytablex.Fields("local")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).local = ""

            End If

            If mytablex.Fields("tipo") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).tipo = mytablex.Fields("tipo")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).tipo = ""

            End If

            If mytablex.Fields("serie") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).serie = mytablex.Fields("serie")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).serie = ""

            End If

            If mytablex.Fields("numero") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).Numero = mytablex.Fields("numero")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).Numero = ""

            End If

            If mytablex.Fields("fecha") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).fecha = mytablex.Fields("fecha")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).fecha = ""

            End If

            If mytablex.Fields("fechae") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).fechae = mytablex.Fields("fechae")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).fechae = ""

            End If

            If mytablex.Fields("hora") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).hora = mytablex.Fields("hora")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).hora = ""

            End If

            If mytablex.Fields("tipoclie") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).tipoclie = mytablex.Fields("tipoclie")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).tipoclie = ""

            End If

            If mytablex.Fields("codigo") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).codigo = mytablex.Fields("codigo")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).codigo = ""

            End If

            If mytablex.Fields("nombre") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).nombre = mytablex.Fields("nombre")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).nombre = ""

            End If

            If mytablex.Fields("moneda") = "S" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).moneda = "Soles"
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).moneda = "Dolares"

            End If

            If mytablex.Fields("total") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).total = mytablex.Fields("total")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).total = 0

            End If

            If mytablex.Fields("acuenta") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).acuenta = mytablex.Fields("acuenta")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).acuenta = 0

            End If

            If mytablex.Fields("adetotal") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).adetotal = mytablex.Fields("adetotal")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).adetotal = 0

            End If

            If mytablex.Fields("bodega") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).bodega = mytablex.Fields("bodega")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).bodega = ""

            End If

            If mytablex.Fields("bodegaf") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).bodegaf = mytablex.Fields("bodegaf")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).bodegaf = ""

            End If

            If mytablex.Fields("localf") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).localf = mytablex.Fields("localf")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).localf = ""

            End If

            If mytablex.Fields("nro_items") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).nro_items = mytablex.Fields("nro_items")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).nro_items = ""

            End If

            If mytablex.Fields("usuario") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).usuario = mytablex.Fields("usuario")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).usuario = ""

            End If

            If mytablex.Fields("placa") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).placa = mytablex.Fields("placa")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).placa = ""

            End If

            If mytablex.Fields("caja") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).caja = mytablex.Fields("caja")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).caja = ""

            End If

            If mytablex.Fields("turno") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).turno = mytablex.Fields("turno")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).turno = ""

            End If

            If mytablex.Fields("vendedor") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).vendedor = mytablex.Fields("vendedor")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).vendedor = ""

            End If

            If mytablex.Fields("observa") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).observa = mytablex.Fields("observa")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).observa = ""

            End If

            If mytablex.Fields("acu") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).acu = mytablex.Fields("acu")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).acu = ""

            End If

            If mytablex.Fields("servicio") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).servicio = mytablex.Fields("servicio")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).servicio = 0

            End If

            ' my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).local1 = mytablex.Fields("local1")
            If mytablex.Fields("tipo1") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).tipo1 = mytablex.Fields("tipo1")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).tipo1 = 0

            End If

            If mytablex.Fields("serie1") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).serie1 = mytablex.Fields("serie1")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).serie1 = 0

            End If

            If mytablex.Fields("numero1") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).numero1 = mytablex.Fields("numero1")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).numero1 = 0

            End If

            If mytablex.Fields("retipo1") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).retipo1 = mytablex.Fields("retipo1")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).retipo1 = 0

            End If

            If mytablex.Fields("renumero3") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).renumero3 = mytablex.Fields("renumero3")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).renumero3 = 0

            End If

            If mytablex.Fields("renumero1") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).renumero1 = mytablex.Fields("renumero1")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).renumero1 = 0

            End If

            If mytablex.Fields("renumero2") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).renumero2 = mytablex.Fields("renumero2")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).renumero2 = 0

            End If

            If mytablex.Fields("neto") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).neto = mytablex.Fields("neto")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).neto = 0

            End If

            If mytablex.Fields("descuento") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).descuento = mytablex.Fields("descuento")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).descuento = 0

            End If

            If mytablex.Fields("subtotal") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).subtotal = mytablex.Fields("subtotal")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).subtotal = 0

            End If

            If mytablex.Fields("impuesto") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).impuesto = mytablex.Fields("impuesto")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).impuesto = 0

            End If

            If mytablex.Fields("tipoimp") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).tipoimp = mytablex.Fields("tipoimp")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).tipoimp = ""

            End If

            If mytablex.Fields("acu1") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).acu1 = mytablex.Fields("acu1")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).acu1 = ""

            End If

            'aqui no visualiza nos datos pero lo considera
            If mytablex.Fields("servicioco") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).servicioco = mytablex.Fields("servicioco")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).servicioco = ""

            End If

            If mytablex.Fields("percepcion") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).percepcion = mytablex.Fields("percepcion")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).percepcion = ""

            End If

            'inicio 16/08/2017 pll
            If mytablex.Fields("DESCRIPCIO") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).descriTipo = mytablex.Fields("DESCRIPCIO")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).descriTipo = ""

            End If

            'fin 16/08/2017 pll
            'inicio 07/09/2017 pll
            If mytablex.Fields("fpago") <> "" Then
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).fpago = mytablex.Fields("fpago")
            Else
                my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).fpago = ""

            End If

            'fin 07/09/2017 pll
            'inicio 0202/2018 pll
            If acu = "V" Then
                If mytablex.Fields("E_SUNAT") <> "" Then
                    my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).esunat = mytablex.Fields("E_SUNAT")
                Else
                    my_struc_deta_cotizacion(UBound(my_struc_deta_cotizacion)).esunat = ""

                End If

            End If

            'fin 0202/2018 pll
            k = k + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

    Exit Sub
cmd921_err:
    MsgBox "aviso en sql_cabeza   " & error$, 48, "Aviso"
    Exit Sub

End Sub

Public Sub nueva_cotizacion(opcion1 As String, _
                            my_local1 As String, _
                            my_struc_nueva_cotizacion() As struc_nueva_cotizacion, _
                            salida As Boolean, _
                            k As Integer)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    ReDim my_struc_nueva_cotizacion(0)

    If opcion1 = "1" Then
        mysql = "select Nombre,Codigo from clientes " & Chr$(10)
    Else
        mysql = "select Nombre,Codigo from clientes where " & Combo1 & " like '" & buffer & "%'" & Chr$(10)

    End If

    If opcion1 = "6100" Then
        mysql = "select Nombre,Codigo from tlocal " & Chr$(10)
    Else
        mysql = "select Nombre,Codigo from tlocal where " & Combo1 & " like '" & buffer & "%'" & Chr$(10)

    End If

    'inicio para el numero de almacenes
    If opcion1 = "7" Then
        mysql = "select Nombre,Codigo from Bodega where local='" & my_local1 & "'" & Chr$(10)

    End If

    'fin **

    If opcion1 = "2" Then
        mysql = "select Producto,Descripcio,Unidad as Und," & Chr$(10)
        mysql = mysql & "Factor as Fac,Precio,Cantidad as Cant,Total,Local," & Chr$(10)
        mysql = mysql & "Deslipo as Dscto from  " & dgusuariog & " " & Chr$(10)
        mysql = mysql & "where local='" & "" & rexplorap.Fields("local") & "' " & Chr$(10)
        mysql = mysql & "and serie='" & "" & rexplorap.Fields("serie") & "' " & Chr$(10)
        mysql = mysql & "and numero='" & "" & rexplorap.Fields("numero") & "'" & Chr$(10)

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
                ReDim Preserve my_struc_nueva_cotizacion(UBound(my_struc_nueva_cotizacion) + 1)

            End If

            my_struc_nueva_cotizacion(UBound(my_struc_nueva_cotizacion)).nombre = mytablex.Fields("nombre")
            k = k + 1
            mytablex.MoveNext
        Loop

    End If

End Sub

Public Function carga_deta_Mcotizacion(my_grid As MSFlexGrid, _
                                       my_struc_deta_cotizacion() As struc_deta_cotizacion, _
                                       k As Integer, _
                                       my_total As Double, _
                                       my_neto As Double, _
                                       my_subtotal As Double, _
                                       my_impuesto As Double, _
                                       my_lvendido As Integer, _
                                       my_anulado As Integer, _
                                       my_tota_anulado As Double, _
                                       my_servicioco As Integer, _
                                       my_percepcion As Integer, _
                                       my_total_agravados As Double, _
                                       my_nume_agravados As Integer)

    For I = 0 To k - 1
        c = c + 1

        my_grid.AddItem ""
        my_grid.Row = my_grid.rows - 1

        my_grid.Col = N_cols_dflag
        my_grid.Text = my_struc_deta_cotizacion(I).dflag

        my_grid.Col = N_cols_yausado
        my_grid.Text = my_struc_deta_cotizacion(I).yausado

        my_grid.Col = N_cols_estado
        my_grid.Text = my_struc_deta_cotizacion(I).estado

        If my_struc_deta_cotizacion(I).estado = "2" Then
            my_lvendido = my_lvendido + 1

        End If

        If my_struc_deta_cotizacion(I).estado = "1" Then
            my_anulado = my_anulado + 1
            my_tota_anulado = my_tota_anulado + my_struc_deta_cotizacion(I).total

        End If

        If my_struc_deta_cotizacion(I).estado = "0" Then
            my_nume_agravados = my_nume_agravados + 1
            my_total_agravados = my_total_agravados + my_struc_deta_cotizacion(I).estado

        End If

        my_grid.Col = N_cols_local
        my_grid.Text = my_struc_deta_cotizacion(I).local

        my_grid.Col = N_cols_tipo
        my_grid.Text = my_struc_deta_cotizacion(I).tipo

        my_grid.Col = N_cols_serie
        my_grid.Text = my_struc_deta_cotizacion(I).serie

        my_grid.Col = N_cols_numero
        my_grid.Text = my_struc_deta_cotizacion(I).Numero

        my_grid.Col = N_cols_fecha
        my_grid.Text = my_struc_deta_cotizacion(I).fecha

        my_grid.Col = N_cols_fechae
        my_grid.Text = my_struc_deta_cotizacion(I).fechae

        my_grid.Col = N_cols_hora
        my_grid.Text = my_struc_deta_cotizacion(I).hora

        my_grid.Col = N_cols_tipoclie
        my_grid.Text = my_struc_deta_cotizacion(I).tipoclie

        my_grid.Col = N_cols_codigo
        my_grid.Text = my_struc_deta_cotizacion(I).codigo

        my_grid.Col = N_cols_nombre
        my_grid.Text = my_struc_deta_cotizacion(I).nombre

        my_grid.Col = N_cols_moneda
        my_grid.Text = my_struc_deta_cotizacion(I).moneda

        my_grid.Col = N_cols_total
        my_grid.Text = my_struc_deta_cotizacion(I).total
        my_total = my_total + my_struc_deta_cotizacion(I).total

        my_grid.Col = N_cols_acuenta
        my_grid.Text = my_struc_deta_cotizacion(I).acuenta

        my_grid.Col = N_cols_adetotal
        my_grid.Text = my_struc_deta_cotizacion(I).adetotal

        my_grid.Col = N_cols_bodega
        my_grid.Text = my_struc_deta_cotizacion(I).bodega

        my_grid.Col = N_cols_bodegaf
        my_grid.Text = my_struc_deta_cotizacion(I).bodegaf

        my_grid.Col = N_cols_localf
        my_grid.Text = my_struc_deta_cotizacion(I).localf

        my_grid.Col = N_cols_nro_items
        my_grid.Text = my_struc_deta_cotizacion(I).nro_items

        my_grid.Col = N_cols_usuario
        my_grid.Text = my_struc_deta_cotizacion(I).usuario

        my_grid.Col = N_cols_placa
        my_grid.Text = my_struc_deta_cotizacion(I).placa

        my_grid.Col = N_cols_caja
        my_grid.Text = my_struc_deta_cotizacion(I).caja

        my_grid.Col = N_cols_turno
        my_grid.Text = my_struc_deta_cotizacion(I).turno

        my_grid.Col = N_cols_vendedor
        my_grid.Text = my_struc_deta_cotizacion(I).vendedor

        my_grid.Col = N_cols_observa
        my_grid.Text = my_struc_deta_cotizacion(I).observa

        my_grid.Col = N_cols_acu
        my_grid.Text = my_struc_deta_cotizacion(I).acu

        my_grid.Col = N_cols_servicio
        my_grid.Text = my_struc_deta_cotizacion(I).servicio

        my_grid.Col = N_cols_retipo1
        my_grid.Text = my_struc_deta_cotizacion(I).retipo1

        my_grid.Col = N_cols_renumero3
        my_grid.Text = my_struc_deta_cotizacion(I).renumero3

        my_grid.Col = N_cols_renumero1
        my_grid.Text = my_struc_deta_cotizacion(I).renumero1

        my_grid.Col = N_cols_renumero1
        my_grid.Text = my_struc_deta_cotizacion(I).renumero2

        my_grid.Col = N_cols_neto
        my_grid.Text = my_struc_deta_cotizacion(I).neto
        my_neto = my_neto + my_struc_deta_cotizacion(I).neto

        my_grid.Col = N_cols_descuento
        my_grid.Text = my_struc_deta_cotizacion(I).descuento

        my_grid.Col = N_cols_subtotal
        my_grid.Text = my_struc_deta_cotizacion(I).subtotal
        my_subtotal = my_subtotal + my_struc_deta_cotizacion(I).subtotal

        my_grid.Col = N_cols_impuesto
        my_grid.Text = my_struc_deta_cotizacion(I).impuesto
        my_impuesto = my_impuesto + my_struc_deta_cotizacion(I).impuesto

        my_grid.Col = N_cols_tipoimp
        my_grid.Text = my_struc_deta_cotizacion(I).tipoimp

        my_grid.Col = N_cols_acu1
        my_grid.Text = my_struc_deta_cotizacion(I).acu1

        'datos que no son visualizados pero considerados pll
        If my_struc_deta_cotizacion(I).servicioco = "" Then
            my_servicioco = my_servicioco + 0
        Else
            my_servicioco = my_servicioco + my_struc_deta_cotizacion(I).servicioco

        End If

        If my_struc_deta_cotizacion(I).percepcion = "" Then
            my_percepcion = my_percepcion + 0
        Else
            my_percepcion = my_percepcion + my_struc_deta_cotizacion(I).percepcion

        End If

        'aqui no visualiza los datos pero son necesarios para elaborar la modificacion pll
        my_grid.Col = N_cols_tipo1
        my_grid.Text = my_struc_deta_cotizacion(I).tipo1

        my_grid.Col = N_cols_serie1
        my_grid.Text = my_struc_deta_cotizacion(I).serie1

        my_grid.Col = N_numero1
        my_grid.Text = my_struc_deta_cotizacion(I).numero1

        my_grid.Col = N_cols_renumero1
        my_grid.Text = my_struc_deta_cotizacion(I).renumero1

        my_grid.Col = N_cols_renumero2
        my_grid.Text = my_struc_deta_cotizacion(I).renumero2

        my_grid.Col = N_cols_renumero3
        my_grid.Text = my_struc_deta_cotizacion(I).renumero3

    Next I

    my_grid.Row = 1
    my_grid.Col = 0

    If my_grid.Text = "" Then
        my_grid.RemoveItem 1

    End If

    Exit Function

End Function

Public Function busca_codigo_local(my_nombre_local As String, my_local1 As String)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo busca_codigo_local

    mysql = "select codigo from tlocal" & Chr$(10)
    mysql = mysql & "where Nombre='" & my_nombre_local & "'" & Chr$(10)

    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF Then
        Exit Function
    Else
        my_local1 = mytablex.Fields("codigo")

    End If

    Exit Function
busca_codigo_local:
    MsgBox "bisca codigo local" & error$, 16, "Aviso"
    Exit Function

End Function

Public Sub Importacion_cotizacion(my_estado As String, _
                                  my_local As String, _
                                  my_tipo As String, _
                                  my_serie As String, _
                                  my_numero As String, _
                                  my_acu As String)
           
    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    mysql = ""
    mysql = "update  " & dgusuariog & "" & Chr$(10)

    'inicio 12/02/2018 pll
    If my_estado = "3" Then
        mysql = mysql & "set estado='3'" & Chr$(10)
    Else
        mysql = mysql & "set estado='1'" & Chr$(10)

    End If

    'fin 12/02/2018 pll
    mysql = mysql & "where  local='" & my_local & "'" & Chr$(10)
    mysql = mysql & "and tipo='" & "" & my_tipo & "' " & Chr$(10)
    mysql = mysql & "and serie='" & "" & my_serie & "'" & Chr$(10)
    mysql = mysql & "and numero='" & "" & my_numero & "'" & Chr$(10)
    mysql = mysql & "and  acu='" & "" & my_acu & "'" & Chr$(10)

    If my_estado = "3" Then
        mysql = mysql & "and estado='1'" & Chr$(10)

    End If

    cn.Execute (mysql)

    mysql = ""
    mysql = "update  fpagov  " & Chr$(10)

    'inicio 12/02/2018 pll
    If my_estado = "3" Then
        mysql = mysql & "set estado='3'" & Chr$(10)
    Else
        mysql = mysql & "set estado='1'" & Chr$(10)

    End If

    'fin 12/02/2018 pll
    mysql = mysql & "where local='" & my_local & "'" & Chr$(10)
    mysql = mysql & "and tipo='" & "" & my_tipo & "' " & Chr$(10)
    mysql = mysql & "and serie='" & "" & my_serie & "'" & Chr$(10)
    mysql = mysql & "and  numero='" & "" & my_numero & "'" & Chr$(10)
    mysql = mysql & "and  acu='" & "" & my_acu & "'" & Chr$(10)

    If my_estado = "3" Then
        mysql = mysql & "and estado='1'" & Chr$(10)

    End If

    cn.Execute (mysql)

    mysql = ""
    mysql = "update " & cgusuario & Chr$(10)

    'inicio 12/02/2018 pll
    If my_estado = "3" Then
        mysql = mysql & "set estado='3'" & Chr$(10)
    Else
        mysql = mysql & "set estado='1'" & Chr$(10)

    End If

    'fin 12/02/2018 pll
    mysql = mysql & "where  local='" & my_local & "'" & Chr$(10)
    mysql = mysql & "and tipo='" & "" & my_tipo & "' " & Chr$(10)
    mysql = mysql & "and serie='" & "" & my_serie & "'" & Chr$(10)
    mysql = mysql & "and numero='" & "" & my_numero & "'" & Chr$(10)
    mysql = mysql & "and acu='" & "" & my_acu & "'" & Chr$(10)

    If my_estado = "3" Then
        mysql = mysql & "and estado='1'" & Chr$(10)

    End If

    cn.Execute (mysql)

End Sub

Public Sub noImportacion_cotizacion(my_estado As String, _
                                    my_local As String, _
                                    my_tipo As String, _
                                    my_serie As String, _
                                    my_numero As String, _
                                    dgusuariog As String, _
                                    my_acu As String)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    mysql = "update  " & gastofactura & "  & Chr$(10)"
    mysql = mysql & "set estado='" & my_estado & "'" & Chr$(10)
    mysql = mysql & "where  local='" & my_local & "'" & Chr$(10)
    mysql = mysql & "and tipo='" & "" & my_tipo & "' " & Chr$(10)
    mysql = mysql & "and serie='" & "" & my_serie & "'" & Chr$(10)
    mysql = mysql & "and  numero='" & "" & my_numero & "'" & Chr$(10)
    mysql = mysql & "and  acu='" & "" & my_acu & "'" & Chr$(10)
    cn.Execute (mysql)

    mysql = ""
    mysql = "update  fpagov  " & Chr$(10)
    mysql = mysql & "set estado='" & buf & "'" & Chr$(10)
    mysql = mysql & "where  local='" & "" & my_local & "' " & Chr$(10)
    mysql = mysql & "and tipo='" & "" & my_tipo & "' " & Chr$(10)
    mysql = mysql & "and serie='" & "" & my_serie & "' " & Chr$(10)
    mysql = mysql & "and  numero='" & "" & my_numero & "'" & Chr$(10)
    mysql = mysql & "and  acu='" & "" & my_acu & "'" & Chr$(10)
    cn.Execute (mysql)

    mysql = ""
    mysql = "update  " & cgusuario & "" & Chr$(10)
    mysql = mysql & "set estado='" & my_estado & "'" & Chr$(10)
    mysql = mysql & "where  local='" & "" & my_local & "'" & Chr$(10)
    mysql = mysql & "and tipo='" & "" & my_tipo & "'" & Chr$(10)
    mysql = mysql & "and serie='" & "" & my_serie & "'" & Chr$(10)
    mysql = mysql & "and numero='" & "" & my_numero & "'" & Chr$(10)
    mysql = mysql & "and  acu='" & "" & my_acu & "'" & Chr$(10)
    cn.Execute (mysql)

End Sub

Public Sub delete_cotizacion(my_estado As String, _
                             my_local As String, _
                             my_tipo As String, _
                             my_serie As String, _
                             my_numero As String, _
                             dgusuariog As String, _
                             cgusuario As String, _
                             acu As String)
                             
    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    mysql = "DELETE FROM  " & dgusuariog & Chr$(10)
    mysql = mysql & "where  local='" & "" & my_local & "'" & Chr$(10)
    mysql = mysql & "and tipo='" & "" & my_tipo & "'" & Chr$(10)
    mysql = mysql & "and serie='" & "" & my_serie & "'" & Chr$(10)
    mysql = mysql & "and  numero='" & "" & my_numero & "'" & buf1 & Chr$(10)
    mysql = mysql & "and acu='" & "" & my_acu & "'" & Chr$(10)
    cn.Execute (mysql)

    mysql = ""
    mysql = "DELETE FROM  fpagov   " & Chr$(10)
    mysql = mysql & "where  local='" & "" & my_local & "'" & Chr$(10)
    mysql = mysql & "and tipo='" & "" & my_tipo & "'" & Chr$(10)
    mysql = mysql & "and serie='" & "" & my_serie & "'" & Chr$(10)
    mysql = mysql & "and  numero='" & "" & my_numero & "'" & buf1 & Chr$(10)
    mysql = mysql & "and acu='" & "" & my_acu & "'" & Chr$(10)
    cn.Execute (mysql)

    mysql = ""
    mysql = "DELETE FROM  " & cgusuario & Chr$(10)
    mysql = mysql & "where  local='" & "" & my_local & "'" & Chr$(10)
    mysql = mysql & "and tipo='" & "" & my_tipo & "'" & Chr$(10)
    mysql = mysql & "and serie='" & "" & my_serie & "'" & Chr$(10)
    mysql = mysql & "and numero='" & "" & my_numero & "'" & Chr$(10)
    mysql = mysql & "and acu='" & "" & my_acu & "'" & Chr$(10)
    cn.Execute (mysql)

    If acu = "3" Then
        mysql = ""
        mysql = "DELETE FROM  serviciotecnico " & Chr$(10)
        mysql = mysql & "where  local='" & "" & my_local & "' " & Chr$(10)
        mysql = mysql & "and tipo='" & "" & my_tipo & "' " & Chr$(10)
        mysql = mysql & "and serie='" & "" & my_serie & "' " & Chr$(10)
        mysql = mysql & "and  numero='" & "" & my_numero & "'" & Chr$(10)
        mysql = mysql & "and acu='" & "" & my_acu & "'" & Chr$(10)
        cn.Execute (mysql)

    End If

End Sub

'aqui controla lave secreta para el vendedor cotizacion
Public Function valida_clave(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM vendedor where  clave='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        If "" & mytablex.Fields("modificacompra") <> "N" Then
            valida_clave = 1

        End If

    End If

    mytablex.Close

End Function

'verifica recibo cotizacion
Public Function verificar_recibo(my_tabla As String, _
                                 my_local As String, _
                                 my_tipo As String, _
                                 my_serie As String, _
                                 my_numero As String, _
                                 found As String)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    mysql = "SELECT * FROM " & my_tabla & "" & Chr$(10)
    mysql = mysql & "where local1='" & my_local & "'" & Chr$(10)
    mysql = mysql & "and tipo1='" & my_tipo & "'" & Chr$(10)
    mysql = mysql & "and serie1='" & my_serie & "'" & Chr$(10)
    mysql = mysql & "and numero1='" & my_numero & "'" & Chr$(10)

    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        verificar_recibo = 1
        found = 1

    End If

    mytablex.Close
    found = 0

End Function

Public Function valida_flag(buf As String)

    Select Case buf

        Case "Z"
            valida_flag = 3

        Case "T", "A", "B", "C", "D", "G", "E", "F"
            valida_flag = 1

        Case "S", "J", "K", "L", "M", "P", "N", "O"
            valida_flag = 2

    End Select

End Function

Public Function borra_detalle(my_local As String, _
                              my_tipo As String, _
                              my_serie As String, _
                              my_numero As String)

    Dim mytablex As New ADODB.Recordset

    Dim mysql    As String

    mysql = "delete from detalle" & Chr$(10)
    mysql = mysql & "where local='" & my_local & "'" & Chr$(10)
    mysql = mysql & "and tipo='" & my_tipo & "' " & Chr$(10)
    mysql = mysql & "and serie='" & xserie & "' " & Chr$(10)
    mysql = mysql & "and numero='" & xnumero & "'" & Chr$(10)
    cn.Execute (mysql)

End Function

Public Function desmarca_yausado(my_local As String, _
                                 my_tipo As String, _
                                 my_serie As String, _
                                 my_numero As String)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd333_err

    mysql = ""
    mysql = "SELECT * FROM " & cgusuario & " " & Chr$(10)
    mysql = mysql & "where  local='" & my_local & "' " & Chr$(10)
    mysql = mysql & "and tipo='" & my_tipo & "' " & Chr$(10)
    mysql = mysql & "and serie='" & my_serie & "' " & Chr$(10)
    mysql = mysql & "and numero='" & my_numero & "'" & Chr$(10)

    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie1"), "" & mytablex.Fields("numero1"), "0"
        descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie2"), "" & mytablex.Fields("numero2"), "0"
        descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie3"), "" & mytablex.Fields("numero3"), "0"
        descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie4"), "" & mytablex.Fields("numero4"), "0"
        descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie5"), "" & mytablex.Fields("numero5"), "0"
        descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie6"), "" & mytablex.Fields("numero6"), "0"
        descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie7"), "" & mytablex.Fields("numero7"), "0"

    End If

    '------------------------------------- ------------
    mytablex.Close
    Exit Function
cmd333_err:
    MsgBox "Aviso en desmarca ya usado " + error$, 48, "Aviso"
    Exit Function
 
End Function

Public Function descarga_el_uso(my_local As String, _
                                my_tipo1 As String, _
                                my_serie1 As String, _
                                my_numero1 As String, _
                                xsw As String)

    If Len(my_tipo) = 0 Then Exit Function
    If Len(my_serie) = 0 Then Exit Function
    If Len(my_numero) = 0 Then Exit Function

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    mysql = "update " & cgusuario & " " & Chr$(10)
    mysql = mysql & "Set yausado = " & xsw & "" & Chr$(10)
    mysql = mysql & "where local='" & my_local & "'" & Chr$(10)
    mysql = mysql & "and tipo='" & my_tipo & "'" & Chr$(10)
    mysql = mysql & "and serie='" & my_serie & "'" & Chr$(10)
    mysql = mysql & "and numero='" & my_numero & "'" & Chr$(10)
    cn.Execute (mysql)

End Function

Public Function graba_acumulado_clientes(buf As String, _
                                         signo As Double, _
                                         sumador As Double)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    mysql = "SELECT * FROM clientes" & Chr$(10)
    mysql = mysql & "where  codigo='" & buf & "'" & Chr$(10)

    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        sdx = Val("" & mytablex.Fields("pedido")) + signo * sumador
        mytablex.Fields("pedido") = sdx
        mytablex.Update

    End If

    mytablex.Close

End Function

Public Function desgraba_cuentac()

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    '---------- validando si es cuenta corriente

    If valida_flag(acu) = 2 Then   'compras
        mysql = "delete from cuentap" & Chr$(10)
        mysql = mysql & "where local='" & rexplorap.Fields("local") & Chr$(10)
        mysql = mysql & "and tipo='" & rexplorap.Fields("tipo") & "'" & Chr$(10)
        mysql = mysql & "and serie='" & rexplorap.Fields("serie") & "'" & Chr$(10)
        mysql = mysql & "and numero='" & rexplorap.Fields("numero") & "'" & Chr$(10)
        cn.Execute (mysql)

    End If

    If valida_flag(acu) = 1 Then   'ventas
        mysql = "delete from cuentac "
        mysql = mysql & "where local='" & rexplorap.Fields("local") & "'" & Chr$(10)
        mysql = mysql & "and tipo='" & rexplorap.Fields("tipo") & "'" & Chr$(10)
        mysql = mysql & "and serie='" & rexplorap.Fields("serie") & "'" & Chr$(10)
        mysql = mysql & "and numero='" & rexplorap.Fields("numero") & "'" & Chr$(10)
        cn.Execute (mysql)

    End If
 
End Function

Public Sub update_cotiza_convert_docu(gtipo As String, _
                                      gserie As String, _
                                      gnumero As String)

    Dim mysql As String

    mysql = ""
    mysql = "update " & cgusuario & Chr$(10)
    mysql = mysql & "set yausado= '1'," & Chr$(10)
    mysql = mysql & "dflag=null," & Chr$(10)
    mysql = mysql & "tipo1='" & extra_loquesea(gtipo) & "'," & Chr$(10)
    mysql = mysql & "serie1='" & gserie & "'," & Chr$(10)
    mysql = mysql & "numero1='" & gnumero & "'," & Chr$(10)
    mysql = mysql & "acu1=null" & Chr$(10)
    mysql = mysql & "where numero='" & my_numero & "'" & Chr$(10)
    cn.Execute (mysql)

End Sub

Public Function ve_descarga(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM tipo where  tipo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe

        Select Case "" & mytablex.Fields("tipodoc")

            Case "A", "B", "C", "D", "G", "E", "F", "J", "K", "L", "M", "P", "N", "O", "S", "T"
                ve_descarga = 1

        End Select

    End If

    mytablex.Close

End Function

Public Function ini_grid_fpago(my_grid As MSFlexGrid)

    my_grid.Clear
    'inizializzazione Grid
    my_grid.rows = 2
    my_grid.FixedRows = 1
    my_grid.Cols = 10 'aqui se aumenta
    my_grid.FixedCols = 0

    my_grid.Row = 0
    my_grid.Col = N_colsm_local 'local
    my_grid.Text = "Local"
    my_grid.ColWidth(my_grid.Col) = 900
    my_grid.FixedAlignment(my_grid.Col) = 0
    my_grid.ColAlignment(my_grid.Col) = 0

    my_grid.Row = 0
    my_grid.Col = N_colsm_tipo 'tipo
    my_grid.Text = "Tipo"
    my_grid.ColWidth(my_grid.Col) = 900
    my_grid.FixedAlignment(my_grid.Col) = 0
    my_grid.ColAlignment(my_grid.Col) = 0

    my_grid.Row = 0
    my_grid.Col = N_colsm_serie 'serie
    my_grid.Text = "Serie"
    my_grid.ColWidth(my_grid.Col) = 1200
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_colsm_numero 'numero
    my_grid.Text = "Numero"
    my_grid.ColWidth(my_grid.Col) = 1200
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_colsm_fpago 'fpago
    my_grid.Text = "F.Pago"
    my_grid.ColWidth(my_grid.Col) = 900
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_colsm_descripcion 'descripcion
    my_grid.Text = "Descripcion"
    my_grid.ColWidth(my_grid.Col) = 2200
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_colsm_moneda 'moneda
    my_grid.Text = "Moneda"
    my_grid.ColWidth(my_grid.Col) = 1200
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_colsm_total 'total
    my_grid.Text = "Total"
    my_grid.ColWidth(my_grid.Col) = 2200
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_colsm_entrega 'fechae 7
    my_grid.Text = "Entrega"
    my_grid.ColWidth(my_grid.Col) = 1800
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_colsm_saldos 'saldos
    my_grid.Text = "Saldos"
    my_grid.ColWidth(my_grid.Col) = 2200
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    Exit Function

End Function

Public Function carica_fpago(my_grid As MSFlexGrid, _
                             my_struc_fpago() As struc_fpago, _
                             k As Integer)

    For I = 0 To k - 1

        my_grid.AddItem ""
        my_grid.Row = my_grid.rows - 1

        my_grid.Col = N_colsm_local
        my_grid.Text = my_struc_fpago(I).local

        my_grid.Col = N_colsm_tipo
        my_grid.Text = my_struc_fpago(I).tipo

        my_grid.Col = N_colsm_serie
        my_grid.Text = my_struc_fpago(I).serie

        my_grid.Col = N_colsm_numero
        my_grid.Text = my_struc_fpago(I).Numero

        my_grid.Col = N_colsm_fpago
        my_grid.Text = my_struc_fpago(I).fpago

        my_grid.Col = N_colsm_descripcion
        my_grid.Text = my_struc_fpago(I).descripcion

        my_grid.Col = N_colsm_moneda
        my_grid.Text = my_struc_fpago(I).moneda

        my_grid.Col = N_colsm_total
        my_grid.Text = my_struc_fpago(I).total

        my_grid.Col = N_colsm_entrega
        my_grid.Text = my_struc_fpago(I).entrega

        my_grid.Col = N_colsm_saldos
        my_grid.Text = my_struc_fpago(I).saldos

    Next I

    my_grid.Row = 1
    my_grid.Col = 0

    If my_grid.Text = "" Then
        my_grid.RemoveItem 1

    End If

    Exit Function

End Function

Public Sub busca_de_tipo(my_local As String, _
                         my_serie As String, _
                         my_numero As String, _
                         my_tipo1 As String, _
                         my_paridad As String, _
                         my_observa As String, _
                         my_dias As String, _
                         my_transporte As String)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    mysql = ""
    mysql = "select distinct tipo1,tipo,paridad,observa,dias,transporte" & Chr$(10)
    mysql = mysql & " from " & cgusuario & Chr$(10)
    mysql = mysql & "where  local='" & "" & my_local & "'" & Chr$(10)
    mysql = mysql & "and numero= '" & my_numero & "'" & Chr$(10)
    mysql = mysql & "and serie= '" & my_serie & "'" & Chr$(10)
 
    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        my_tipo1 = mytablex.Fields("tipo")
        my_paridad = mytablex.Fields("paridad")
        my_observa = mytablex.Fields("observa")
        my_dias = mytablex.Fields("dias")
        my_transporte = mytablex.Fields("transporte")

    End If

    mytablex.Close

End Sub

Public Function busca_codigo_bodega(my_nombre_bodega As String, my_bodegaf As String)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo busca_codigo_bodega

    mysql = "select codigo from bodega" & Chr$(10)
    mysql = mysql & "where Nombre='" & my_nombre_bodega & "'" & Chr$(10)

    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF Then
        Exit Function
    Else
        my_bodegaf = mytablex.Fields("codigo")

    End If

    Exit Function
busca_codigo_bodega:
    MsgBox "bisca codigo bodega" & error$, 16, "Aviso"
    Exit Function

End Function

'inicio 14/11/2017 pll **
Public Function busca_descri_bodega(my_cod_bodega As String, my_describodega As String)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo busca_codigo_bodega

    mysql = "select nombre from bodega" & Chr$(10)
    mysql = mysql & "where codigo='" & my_cod_bodega & "'" & Chr$(10)

    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF Then
        Exit Function
    Else
        my_describodega = mytablex.Fields("nombre")

    End If

    Exit Function
busca_codigo_bodega:
    MsgBox "bisca codigo bodega" & error$, 16, "Aviso"
    Exit Function

End Function

' Testing Proyecto Facturacion Electronica 05/04/2018
' Testing Proyecto Facturacion Electronica
Public Function obtiene_EstadoSunat(cgusuario As String, _
                                    my_local As String, _
                                    my_tipo As String, _
                                    my_serie As String, _
                                    my_numero As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT estado_sunat FROM  " & cgusuario & "  where  local='" & my_local & "' and tipo='" & my_tipo & "' and serie='" & my_serie & "' and numero='" & my_numero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
          
        If IsNull(mytablex.Fields("estado_sunat")) Then
            obtiene_EstadoSunat = 0 'Pendiente
        Else
            obtiene_EstadoSunat = mytablex.Fields("estado_sunat") '3 es ya dado de baja

        End If
     
    Else
        obtiene_EstadoSunat = "NO EXISTE"

    End If

    mytablex.Close

End Function

' Testing Proyecto Facturacion Electronica
