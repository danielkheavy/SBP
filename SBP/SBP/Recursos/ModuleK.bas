Attribute VB_Name = "Module25"
''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
Option Explicit

Function ObtieneStockInicial(xlocal1 As String, _
                             xproducto As String, _
                             xbodega As String, _
                             xfechai As String, _
                             xfechaf As String)

    Dim buf As String

    buf = " delete from talmacen " '  where producto like '" & xproducto & "'"
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
    buf = buf & "                                             detalle.acu = 'G' or detalle.acu='E' or detalle.acu='T' or detalle.acu='N')"
    buf = buf & "                       GROUP BY detalle.local,detalle.producto,detalle.bodega) "

    cn.Execute (buf)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select  sum(talmacen.saldo) as saldo from talmacen ", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        ObtieneStockInicial = mytablex.Fields("saldo")

        If IsNull(mytablex.Fields("saldo")) Then ObtieneStockInicial = "0"

    End If

    ObtieneStockInicial = ObtieneStockInicial

End Function

''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
