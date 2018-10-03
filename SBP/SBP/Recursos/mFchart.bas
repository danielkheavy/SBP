Attribute VB_Name = "mFchart"
'inicio 27/06/2017 pll
'para carge aumatico las combo
'en el excel chart
Type struct_viejo_analisis
  mes                                   As String
  moneda                                As String
  total                                 As Double
End Type

'Type struct_moneda
'  moneda                                As String
'End Type

Type struct_TipoDocumento
  Descripcion                           As String
End Type

Type struct_servicio
  descripcio                            As String
End Type

Type struct_moneda
  moneda                                As String
End Type

Type struct_caja
  caja                                  As String
End Type

Type struct_vendedor
  nombre                                As String
End Type

Type struct_turno
   turno                                As String
   horai                                As String
   horaf                                As String
End Type

Type struct_familia
   familia                              As String
End Type

Public carga_viejo_analisis()           As struct_viejo_analisis
Public carga_tipoDocumento()            As struct_TipoDocumento
Public carga_servicio()                 As struct_servicio
Public carga_moneda()                   As struct_moneda
Public carga_caja()                     As struct_caja
Public carga_vendedor()                 As struct_vendedor
Public carga_turno()                    As struct_turno
Public carga_familia()                  As struct_familia
'Public carga_moneda()                   As struct_moneda

Public my_moneda                        As String
Public my_caja                          As String
Public my_turno                         As String
Public my_vendedor                      As String
'Public my_servicio                      As String
Public my_familia                       As String
Public my_tgrafico                      As String * 1
'inicio 08/08/2017 pll
Type struct_reportG
    fecha                               As String
    cajaVentas                          As String
    TotalVenta                          As Double
    cajaCompras                         As String
    TotalCompras                        As Double
End Type
Global my_struct_reportG()       As struct_reportG
'fin 08/08/2017 pll
Public Sub l_TipoDocumento(carga_tipoDocumento() As struct_TipoDocumento, salida As Boolean, k As Integer)

Dim mysql                     As String
Dim mytablef                  As New ADODB.Recordset


ReDim carga_tipoDocumento(0)

  mysql = "SELECT  descripcio from tipo" & Chr$(10)
  mysql = mysql & "order by tipo asc" & Chr$(10)
  
  mytablef.Open mysql, cn, adOpenStatic, adLockOptimistic
  If mytablef.EOF Then
     salida = True
     Exit Sub
  Else
    mytablef.MoveFirst
    k = 0
    Do Until mytablef.EOF
     If k > 0 Then
        ReDim Preserve carga_tipoDocumento(UBound(carga_tipoDocumento) + 1)
     End If
     carga_tipoDocumento(UBound(carga_tipoDocumento)).Descripcion = RTrim(mytablef.Fields("descripcio"))
     k = k + 1
     mytablef.MoveNext
    Loop
    salida = False
    Exit Sub
  End If
  
End Sub
Public Sub l_servicio(carga_servicio() As struct_servicio, salida As Boolean, k As Integer)

Dim mysql                     As String
Dim mytablef                  As New ADODB.Recordset


ReDim carga_servicio(0)

  mysql = "SELECT  descripcio from servicio" & Chr$(10)
  mysql = mysql & "order by descripcio desc" & Chr$(10)
  
  mytablef.Open mysql, cn, adOpenStatic, adLockOptimistic
  If mytablef.EOF Then
     salida = True
     Exit Sub
  Else
    mytablef.MoveFirst
    k = 0
    Do Until mytablef.EOF
     If k > 0 Then
        ReDim Preserve carga_servicio(UBound(carga_servicio) + 1)
     End If
     carga_servicio(UBound(carga_servicio)).descripcio = RTrim(mytablef.Fields("descripcio"))
     k = k + 1
     mytablef.MoveNext
    Loop
    salida = False
    Exit Sub
  End If
  
End Sub
Public Sub l_moneda(carga_moneda() As struct_moneda, salida As Boolean, k As Integer)

Dim mysql                     As String
Dim mytablef                  As New ADODB.Recordset


ReDim carga_moneda(0)

  mysql = "SELECT DISTINCT MONEDA =" & Chr$(10)
  mysql = mysql & "CASE MONEDA WHEN 'S' THEN 'Soles'" & Chr$(10)
  mysql = mysql & "WHEN 'D' THEN 'Dolares' End" & Chr$(10)
  mysql = mysql & "FROM FPAGO" & Chr$(10)
  
  mytablef.Open mysql, cn, adOpenStatic, adLockOptimistic
  If mytablef.EOF Then
     salida = True
     Exit Sub
  Else
    mytablef.MoveFirst
    k = 0
    Do Until mytablef.EOF
     If k > 0 Then
        ReDim Preserve carga_moneda(UBound(carga_moneda) + 1)
     End If
     If mytablef.Fields("moneda") <> "" Then
       carga_moneda(UBound(carga_moneda)).moneda = mytablef.Fields("moneda")
      Else
      carga_moneda(UBound(carga_moneda)).moneda = ""
     End If
     k = k + 1
     mytablef.MoveNext
    Loop
    salida = False
    Exit Sub
  End If
End Sub
Public Sub l_caja(carga_caja() As struct_caja, salida As Boolean, k As Integer)

Dim mysql                     As String
Dim mytablef                  As New ADODB.Recordset


ReDim carga_caja(0)

  mysql = "SELECT caja" & Chr$(10)
  mysql = mysql & "FROM parameca" & Chr$(10)
  mysql = mysql & "order by caja asc" & Chr$(10)
  
  mytablef.Open mysql, cn, adOpenStatic, adLockOptimistic
  If mytablef.EOF Then
     salida = True
     Exit Sub
  Else
    mytablef.MoveFirst
    k = 0
    Do Until mytablef.EOF
     If k > 0 Then
        ReDim Preserve carga_caja(UBound(carga_caja) + 1)
     End If
     If mytablef.Fields("caja") <> "" Then
      carga_caja(UBound(carga_caja)).caja = mytablef.Fields("caja")
     Else
      carga_caja(UBound(carga_caja)).caja = ""
     End If
     k = k + 1
     mytablef.MoveNext
    Loop
    salida = False
    Exit Sub
  End If
End Sub
Public Sub l_vendedor(carga_vendedor() As struct_vendedor, salida As Boolean, k As Integer)

Dim mysql                     As String
Dim mytablef                  As New ADODB.Recordset


ReDim carga_vendedor(0)

' mysql = "SELECT nombre" & Chr$(10)
' mysql = mysql & "FROM vendedor" & Chr$(10)
' mysql = mysql & "where estado ='ACTIVO'" & Chr$(10)
' mysql = mysql & "order by caja asc" & Chr$(10)
  
mysql = "select distinct  vendedor" & Chr$(10)
mysql = mysql & "From factura" & Chr$(10)
mysql = mysql & "Where Vendedor Is Not Null" & Chr$(10)
mysql = mysql & "and vendedor <>''" & Chr$(10)

  mytablef.Open mysql, cn, adOpenStatic, adLockOptimistic
  If mytablef.EOF Then
     salida = True
     Exit Sub
  Else
    mytablef.MoveFirst
    k = 0
    Do Until mytablef.EOF
     If k > 0 Then
        ReDim Preserve carga_vendedor(UBound(carga_vendedor) + 1)
     End If
     If mytablef.Fields("vendedor") <> "" Then
      carga_vendedor(UBound(carga_vendedor)).nombre = mytablef.Fields("vendedor")
     Else
      carga_vendedor(UBound(carga_vendedor)).nombre = ""
     End If
     k = k + 1
     mytablef.MoveNext
    Loop
    salida = False
    Exit Sub
  End If
End Sub
Public Sub l_turno(carga_turno() As struct_turno, salida As Boolean, k As Integer)

Dim mysql                     As String
Dim mytablef                  As New ADODB.Recordset


ReDim carga_turno(0)

 mysql = "SELECT turno,horai,horaf" & Chr$(10)
 mysql = mysql & "FROM TURNO" & Chr$(10)
 mysql = mysql & "order by TURNO asc" & Chr$(10)
  
  mytablef.Open mysql, cn, adOpenStatic, adLockOptimistic
  If mytablef.EOF Then
     salida = True
     Exit Sub
  Else
    mytablef.MoveFirst
    k = 0
    Do Until mytablef.EOF
     If k > 0 Then
        ReDim Preserve carga_turno(UBound(carga_turno) + 1)
     End If
     If mytablef.Fields("turno") <> "" Then
      carga_turno(UBound(carga_turno)).turno = mytablef.Fields("turno")
     Else
      carga_turno(UBound(carga_turno)).turno = ""
     End If
     If mytablef.Fields("horai") <> "" Then
      carga_turno(UBound(carga_turno)).horai = mytablef.Fields("horai")
     Else
      carga_turno(UBound(carga_turno)).horai = ""
     End If
     If mytablef.Fields("horaf") <> "" Then
      carga_turno(UBound(carga_turno)).horaf = mytablef.Fields("horaf")
     Else
      carga_turno(UBound(carga_turno)).horaf = ""
     End If
     
     k = k + 1
     mytablef.MoveNext
    Loop
    salida = False
    Exit Sub
  End If
End Sub
Public Sub l_familia(carga_familia() As struct_familia, salida As Boolean, k As Integer)

Dim mysql                     As String
Dim mytablef                  As New ADODB.Recordset


ReDim carga_familia(0)

 mysql = "SELECT familia" & Chr$(10)
 mysql = mysql & "FROM FAMILIA" & Chr$(10)
 mysql = mysql & "order by FAMILIA asc" & Chr$(10)
  
  mytablef.Open mysql, cn, adOpenStatic, adLockOptimistic
  If mytablef.EOF Then
     salida = True
     Exit Sub
  Else
    mytablef.MoveFirst
    k = 0
    Do Until mytablef.EOF
     If k > 0 Then
        ReDim Preserve carga_familia(UBound(carga_familia) + 1)
     End If
     If mytablef.Fields("familia") <> "" Then
      carga_familia(UBound(carga_familia)).familia = mytablef.Fields("familia")
     Else
      carga_familia(UBound(carga_familia)).familia = ""
     End If
     k = k + 1
     mytablef.MoveNext
    Loop
    salida = False
    Exit Sub
  End If
End Sub

Public Sub crear_chart(objWorkBook As Excel.Workbook, v As Integer, h As Integer)
Dim objExcel                           As Excel.Application
Dim objWorksheet                       As Excel.Worksheet
Dim chtobj                             As ChartObject
Dim my_final                           As Integer

 On Error GoTo crear_chart
 
 Set objWorksheet = objWorkBook.Worksheets(1)

 
 ir = objWorksheet.Cells(rows.count, 1).End(xlUp).Row
 Set chtRng = Range("A13:E" & ir + v)

 my_final = v - 1
 
 objWorkBook.Charts.Add
 objWorkBook.ActiveChart.SetSourceData Source:=objWorkBook.Sheets("Hoja1").Range("A13:E" & my_final), PlotBy:=xlColumns
 objWorkBook.ActiveChart.SeriesCollection(1).Name = "=""Fecha"""


 With ActiveChart
  .HasTitle = True
  .ChartStyle = 34
  .ChartTitle.Characters.Text = "Analisis Compras Vs.Ventas"
  .Axes(xlCategory, xlPrimary).HasTitle = True
  .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Caja Venta-Total Venta-Caja Compras / Fecha"
  .Axes(xlValue, xlPrimary).HasTitle = True
  .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Cantidades"

 End With
crear_chart:
 Select Case Err.Number
        Case 1004
            Resume Next
        Case Else
           ' MsgBox Err.Number & vbcrlf & Err.Description
    End Select
End Sub
Public Sub detalle_chart_sele(fechai As String, fechaf As String, my_moneda As String, my_caja As String, _
           my_turno As String, my_vendedor As String, my_servicio As String, my_familia As String, _
           k As Integer, salida As Boolean, my_struct_reportG() As struct_reportG)
           
Dim mysql                     As String
Dim mytable                   As New ADODB.Recordset

On Error GoTo detalle_chart_sele

ReDim my_struct_reportG(0)

mysql = "select ventas.fecha as fecha,ventas.CAJA as cajaVentas,ventas.TotalVenta as TotalVenta," & Chr$(10)
mysql = mysql & "compras.CAJA as cajaCompras,compras.TotalCompras as TotalCompras" & Chr$(10)
mysql = mysql & "From" & Chr$(10)
mysql = mysql & "(select de.fecha,de.caja,sum(de.total) as TotalVenta" & Chr$(10)
mysql = mysql & "from detalle de," & Chr$(10)
mysql = mysql & "tipo ti" & Chr$(10)
mysql = mysql & "where de.fecha>='" & Format(fechai, "YYYYMMDD") & "'" & Chr$(10)
mysql = mysql & "and de.fecha<='" & Format(fechaf, "YYYYMMDD") & "'" & Chr$(10)
mysql = mysql & "and ltrim(ti.tipo) = ltrim(de.tipo)" & Chr$(10)
If my_moneda <> "%" Then
   mysql = mysql & "and moneda like '" & my_moneda & "'" & Chr$(10)
End If
If my_servicio <> "%" Then
  mysql = mysql & "and servicio like '" & my_servicio & "'" & Chr$(10)
End If
If my_caja <> "%" Then
  mysql = mysql & "and caja like '" & my_caja & "'" & Chr$(10)
End If
If my_turno <> "%" Then
  mysql = mysql & "and turno like '" & my_turno & "'" & Chr$(10)
End If
If my_vendedor <> "%" Then
  'mysql = mysql & "and vendedor like '" & extra_loquesea(Vendedor) & "'" & Chr$(10)
  mysql = mysql & "and vendedor like '" & vendedor & "'" & Chr$(10)
End If
If my_familia <> "%" Then
   mysql = mysql & "and familia like '" & my_familia & "'" & Chr$(10)
End If
mysql = mysql & "and de.tipo <>'BC'" & Chr$(10)
mysql = mysql & "group by de.fecha,de.caja)ventas," & Chr$(10)
mysql = mysql & "(select de.fecha,de.caja,sum(de.total) as TotalCompras" & Chr$(10)
mysql = mysql & "from detalle de," & Chr$(10)
mysql = mysql & "tipo ti" & Chr$(10)
mysql = mysql & "where de.fecha>='" & Format(fechai, "YYYYMMDD") & "'" & Chr$(10)
mysql = mysql & "and de.fecha<='" & Format(fechaf, "YYYYMMDD") & "'" & Chr$(10)
mysql = mysql & "and ltrim(ti.tipo) = ltrim(de.tipo)" & Chr$(10)
If my_moneda <> "%" Then
   mysql = mysql & "and moneda like '" & my_moneda & "'" & Chr$(10)
End If
If my_servicio <> "%" Then
  mysql = mysql & "and servicio like '" & my_servicio & "'" & Chr$(10)
End If
If my_caja <> "%" Then
  mysql = mysql & "and caja like '" & my_caja & "'" & Chr$(10)
End If
If my_turno <> "%" Then
  mysql = mysql & "and turno like '" & my_turno & "'" & Chr$(10)
End If
If my_vendedor <> "%" Then
  'mysql = mysql & "and vendedor like '" & extra_loquesea(Vendedor) & "'" & Chr$(10)
  mysql = mysql & "and vendedor like '" & vendedor & "'" & Chr$(10)
End If
If my_familia <> "%" Then
   mysql = mysql & "and familia like '" & my_familia & "'" & Chr$(10)
End If
mysql = mysql & "and de.tipo ='BC'" & Chr$(10)
mysql = mysql & "group by de.fecha,de.caja)compras" & Chr$(10)

mytable.Open mysql, cn, adOpenStatic, adLockOptimistic

 If mytable.EOF Then
    salida = False
    Exit Sub
 Else
   salida = True
   mytable.MoveFirst
   k = 0
   Do Until mytable.EOF
     If k > 0 Then
         ReDim Preserve my_struct_reportG(UBound(my_struct_reportG) + 1)
     End If
     If mytable.Fields("fecha") <> "" Then
       my_struct_reportG(UBound(my_struct_reportG)).fecha = mytable.Fields("fecha")
     Else
      my_struct_reportG(UBound(my_struct_reportG)).fecha = ""
     End If
     If mytable.Fields("cajaVentas") <> "" Then
       my_struct_reportG(UBound(my_struct_reportG)).cajaVentas = mytable.Fields("cajaVentas")
     Else
      my_struct_reportG(UBound(my_struct_reportG)).cajaVentas = ""
     End If
     If mytable.Fields("TotalVenta") <> "" Then
      my_struct_reportG(UBound(my_struct_reportG)).TotalVenta = mytable.Fields("TotalVenta")
     Else
       my_struct_reportG(UBound(my_struct_reportG)).TotalVenta = 0
     End If
     If mytable.Fields("cajaCompras") <> "" Then
       my_struct_reportG(UBound(my_struct_reportG)).cajaCompras = mytable.Fields("cajaCompras")
     Else
      my_struct_reportG(UBound(my_struct_reportG)).cajaCompras = ""
     End If
    If mytable.Fields("TotalCompras") <> "" Then
     my_struct_reportG(UBound(my_struct_reportG)).TotalCompras = mytable.Fields("TotalCompras")
    Else
     my_struct_reportG(UBound(my_struct_reportG)).TotalCompras = 0
    End If
     k = k + 1
    mytable.MoveNext
  Loop
 End If
mytable.Close
Exit Sub
detalle_chart_sele:
Select Case Err.Number
        Case 1004
            Resume Next
        Case Else
            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select
End Sub
Public Sub por_documentos(carga_viejo_analisis() As struct_viejo_analisis, my_tipo_analisis As String, my_moneda As String, _
           my_caja As String, my_servicio As String, my_vendedor As String, my_turno As String, _
           my_familia As String, my_fechai As String, my_fechaf As String, my_tipo As String, _
           my_acu As String, my_codigo As String, salida As Boolean, k As Integer, _
           carga_moneda() As struct_moneda)
           
Dim mysql                      As String
Dim mytable                    As New ADODB.Recordset
Dim mytablex                   As New ADODB.Recordset

   swbuf = "sum(total) as xtotal"
   If cantidad = "Cantidad" Then
   swbuf = "count(numero) as xtotal"
   End If
    'aqui debe cuadrar la tabla tipo con la tabla caja
   'If my_busqueda = "Documentos" Then 'aqui usa la tabla tipo
     If my_tipo_analisis = "Caja" Then
     
        ReDim carga_viejo_analisis(0)
        
        mysql = "select d.Caja as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from  factura d," & Chr$(10)
        mysql = mysql & "parameca p," & Chr$(10)
        mysql = mysql & "tipo t," & Chr$(10)
        mysql = mysql & "producto Pr" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.CAJA = p.caja" & Chr$(10)
        mysql = mysql & "and t.tipo = d.tipo" & Chr$(10)
        If my_tipo <> "%" Then
         mysql = mysql & "and d.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and d.codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and pr.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and d.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and p.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and d.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and pr.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by d.Caja ,d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
        If mytable.Fields("mes") <> "" Then
          carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
        Else
          carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
        End If
        If mytable.Fields("moneda") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
        Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
        End If
        If mytable.Fields("xtotal") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
        Else
          carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
        End If
       k = k + 1
       mytable.MoveNext
      Loop
      
      mytable.Close

     End If
  
    End If
   '**aqui es el tipo
   If my_tipo_analisis = "Tipo" Then
     
        ReDim carga_viejo_analisis(0)
        
        mysql = "select d.tipo as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "tipo t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and t.tipo = d.tipo" & Chr$(10)

        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and d.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and d.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by d.tipo ,d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
        If mytable.Fields("mes") <> "" Then
          carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
        Else
          carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
       End If
       If mytable.Fields("moneda") <> "" Then
          carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
       Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
       End If
       If mytable.Fields("xtotal") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
       Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
       End If
       k = k + 1
       mytable.MoveNext
      Loop
      
      mytable.Close

     End If
  
    End If
   '**aqui es el usuario
   If my_tipo_analisis = "Usuario" Then
     
        ReDim carga_viejo_analisis(0)
        
        mysql = "select d.usuario as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "tipo t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and t.tipo = d.tipo" & Chr$(10)

        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and d.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by d.usuario ,d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
        If mytable.Fields("mes") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
        Else
          carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
       End If
       If mytable.Fields("moneda") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
       Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
       End If
       If mytable.Fields("xtotal") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
       End If
       k = k + 1
       mytable.MoveNext
      Loop
      
      mytable.Close

     End If
    End If
    '**aqui es para la bodega
    If my_tipo_analisis = "Bodega" Then
     
        ReDim carga_viejo_analisis(0)
        
        mysql = "select d.bodega as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "tipo t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and t.tipo = d.tipo" & Chr$(10)

        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and d.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by d.bodega ,d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
        If mytable.Fields("mes") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
        Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
        End If
        If mytable.Fields("moneda") <> "" Then
           carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
        Else
          carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
        End If
        If mytable.Fields("xtotal") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
        Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
        End If
       k = k + 1
       mytable.MoveNext
      Loop
      mytable.Close

     End If
    End If
    
    '**aqui es para la codigo
    If my_tipo_analisis = "Codigo" Then
     
        ReDim carga_viejo_analisis(0)
        
        mysql = "select d.codigo as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "tipo t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and t.tipo = d.tipo" & Chr$(10)

        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and t.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by d.codigo,d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
       If mytable.Fields("mes") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
       Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
       End If
       If mytable.Fields("moneda") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
       End If
       If mytable.Fields("xtotal") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
        Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
       End If
       k = k + 1
       mytable.MoveNext
      Loop
      mytable.Close

     End If
    End If
     'aqui es para el vendedor
     If my_tipo_analisis = "Vendedor" Then
     
        ReDim carga_viejo_analisis(0)
        
        mysql = "select d.vendedor as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "tipo t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and t.tipo = d.tipo" & Chr$(10)

        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and d.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by d.vendedor,d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
        If mytable.Fields("mes") <> "" Then
           carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
          
        Else
            carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
        End If
       If mytable.Fields("moneda") <> "" Then
          carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
          
       Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
       End If
       If mytable.Fields("xtotal") <> "" Then
          carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
       Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
       End If
       k = k + 1
       mytable.MoveNext
      Loop
      
      mytable.Close

     End If
    End If
    'aqui es para anual,mensual,semanal,diario,horario

      If my_tipo_analisis = "Anual" Then
        ReDim carga_viejo_analisis(0)
        
        mysql = "select year(d.fecha) as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "tipo t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and t.tipo = d.tipo" & Chr$(10)

        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and d.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by year(d.fecha),d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
       If mytable.Fields("mes") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
       Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
       End If
       If mytable.Fields("moneda") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
       Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
       End If
       If mytable.Fields("xtotal") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
       Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
       End If
       k = k + 1
       mytable.MoveNext
      Loop
      mytable.Close

     End If
    End If
    '** mensual
    If my_tipo_analisis = "Mensual" Then
        ReDim carga_viejo_analisis(0)
        
        mysql = "select month(d.fecha) as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "tipo t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and t.tipo = d.tipo" & Chr$(10)

        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and d.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by month(d.fecha),d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
        If mytable.Fields("mes") <> "" Then
          carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
        Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
        End If
        If mytable.Fields("moneda") <> "" Then
          carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
        Else
          carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
        End If
        If mytable.Fields("xtotal") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
        Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
        End If
       k = k + 1
       mytable.MoveNext
      Loop
      mytable.Close

     End If
    End If
   'semanal
   If my_tipo_analisis = "Semanal" Then
        ReDim carga_viejo_analisis(0)
        
        mysql = "select DATENAME(weekday,d.fecha) as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "tipo t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and t.tipo = d.tipo" & Chr$(10)

        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and d.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by DATENAME(weekday,d.fecha),d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
        If mytable.Fields("mes") <> "" Then
          carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
        Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
        End If
        If mytable.Fields("moneda") <> "" Then
          carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
        Else
          carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
        End If
        If mytable.Fields("xtotal") <> "" Then
          carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
        Else
          carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
        End If
       k = k + 1
       mytable.MoveNext
      Loop

      mytable.Close

     End If
    End If
   'diario
    If my_tipo_analisis = "Diario" Then
        ReDim carga_viejo_analisis(0)
        
        mysql = "select day(d.fecha) as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "tipo t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and t.tipo = d.tipo" & Chr$(10)

        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and d.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by day(d.fecha),d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
        If mytable.Fields("mes") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
        Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
       End If
       If mytable.Fields("moneda") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
       End If
       If mytable.Fields("total") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("total")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
       End If
       k = k + 1
       mytable.MoveNext
      Loop
      mytable.Close

     End If
    End If
   'hora
    If my_tipo_analisis = "Horario" Then
        ReDim carga_viejo_analisis(0)
        
        mysql = "select left(d.hora,2) as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "tipo t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and t.tipo = d.tipo" & Chr$(10)

        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and d.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by left(d.hora,2),d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
       If mytable.Fields("mes") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
       Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
       End If
       If mytable.Fields("moneda") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
       End If
       If mytable.Fields("xtotal") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
       Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
       End If
       k = k + 1
       mytable.MoveNext
      Loop
      mytable.Close

     End If
    End If
  Exit Sub
  
End Sub
Public Sub por_clientes(carga_viejo_analisis() As struct_viejo_analisis, my_tipo_analisis As String, my_moneda As String, _
           my_caja As String, my_servicio As String, my_vendedor As String, my_turno As String, _
           my_familia As String, my_fechai As String, my_fechaf As String, my_tipo As String, _
           my_acu As String, my_codigo As String, salida As Boolean, k As Integer, _
           carga_moneda() As struct_moneda)
           
Dim mysql                     As String
Dim mytable                   As New ADODB.Recordset
Dim mytablex                  As New ADODB.Recordset

   swbuf = "sum(total) as xtotal"
   If cantidad = "Cantidad" Then
   swbuf = "count(numero) as xtotal"
   End If
    'aqui debe cuadrar la tabla tipo con la tabla caja
   'If my_busqueda = "Documentos" Then 'aqui usa la tabla tipo
     If my_tipo_analisis = "Caja" Then
     
        ReDim carga_viejo_analisis(0)
        
        mysql = "select d.TIPOCLIE as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "parameca p," & Chr$(10)
        mysql = mysql & "clientes t," & Chr$(10)
        mysql = mysql & "producto pr" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.TIPOCLIE='C'" & Chr$(10)
        mysql = mysql & "and d.CAJA = p.caja" & Chr$(10)
        'mysql = mysql & "and d.codigo = t.codigo" & Chr$(10)
  
        If my_tipo <> "%" Then
         mysql = mysql & "and d.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and d.codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and pr.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and d.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and p.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and d.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and pr.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by d.TIPOCLIE ,d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
        If mytable.Fields("mes") <> "" Then
          carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
        Else
          carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
        End If
        If mytable.Fields("moneda") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
        Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
        End If
        If mytable.Fields("xtotal") <> "" Then
          carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
        Else
          carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
        End If
       k = k + 1
       mytable.MoveNext
      Loop
       
     mytable.Close

     End If
  
    End If
   '**aqui es el tipo
   If my_tipo_analisis = "Tipo" Then
     
        ReDim carga_viejo_analisis(0)
        
        mysql = "select d.tipo as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from  factura d," & Chr$(10)
        mysql = mysql & "clientes t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.TIPOCLIE='C'" & Chr$(10)
        'mysql = mysql & "and d.codigo = t.codigo" & Chr$(10)
        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and t.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by d.tipo ,d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
       If mytable.Fields("mes") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
       End If
       If mytable.Fields("moneda") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
       End If
       If mytable.Fields("xtotal") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
       End If
       k = k + 1
       mytable.MoveNext
      Loop
      mytable.Close
      
     End If
  
    End If
   '**aqui es el usuario
   If my_tipo_analisis = "Usuario" Then
     
        ReDim carga_viejo_analisis(0)
        
        mysql = "select d.usuario as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "clientes t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.TIPOCLIE='C'" & Chr$(10)
        'mysql = mysql & "and d.codigo = t.codigo" & Chr$(10)

        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and t.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by d.usuario ,d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
        If mytable.Fields("mes") <> "" Then
          carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
        Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
        End If
        If mytable.Fields("moneda") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
        Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
        End If
        If mytable.Fields("xtotal") <> "" Then
          carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
        Else
          carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
        End If
       k = k + 1
       mytable.MoveNext
      Loop
      
      mytable.Close
      Exit Sub
 
     End If
    End If
    '**aqui es para la bodega
    If my_tipo_analisis = "Bodega" Then
     
        ReDim carga_viejo_analisis(0)
        
        mysql = "select d.bodega as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "clientes t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.TIPOCLIE='C'" & Chr$(10)
        'mysql = mysql & "and d.codigo = t.codigo" & Chr$(10)
        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and t.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by d.bodega ,d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
       If mytable.Fields("mes") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
       End If
       If mytable.Fields("moneda") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
       End If
       If mytable.Fields("xtotal") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
       End If
       k = k + 1
       mytable.MoveNext
      Loop
      mytable.Close
     End If
    End If
    
    '**aqui es para la codigo
    If my_tipo_analisis = "Codigo" Then
     
        ReDim carga_viejo_analisis(0)
        
        mysql = "select d.codigo as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "clientes t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.TIPOCLIE='C'" & Chr$(10)
        'mysql = mysql & "and d.codigo = t.codigo" & Chr$(10)
        
        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and t.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by d.codigo,d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
       If mytable.Fields("mes") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
       Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
       End If
       If mytable.Fields("moneda") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
       Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
       End If
       If mytable.Fields("xtotal") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
       End If
       k = k + 1
       mytable.MoveNext
      Loop

      mytable.Close
    
     End If
    End If
     'aqui es para el vendedor
     If my_tipo_analisis = "Vendedor" Then
     
        ReDim carga_viejo_analisis(0)
        
        mysql = "select d.vendedor as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "clientes t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.TIPOCLIE='C'" & Chr$(10)
        'mysql = mysql & "and d.codigo = t.codigo" & Chr$(10)
        
        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and t.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by d.vendedor,d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
        If mytable.Fields("mes") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
        Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
        End If
        If mytable.Fields("moneda") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
        Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
        End If
        If mytable.Fields("xtotal") <> "" Then
          carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
        Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
        End If
       k = k + 1
       mytable.MoveNext
      Loop
      
      mytable.Close

     End If
    End If
    'aqui es para anual,mensual,semanal,diario,horario
     If my_tipo_analisis = "Anual" Then
        ReDim carga_viejo_analisis(0)
        
        mysql = "select year(d.fecha) as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "clientes t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.TIPOCLIE='C'" & Chr$(10)
        'mysql = mysql & "and d.codigo = t.codigo" & Chr$(10)
        
        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and t.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by year(d.fecha),d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
        If mytable.Fields("mes") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
        Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
        End If
        If mytable.Fields("moneda") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
       End If
       If mytable.Fields("xtotal") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
       Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
       End If
       k = k + 1
       mytable.MoveNext
      Loop
      mytable.Close
  
     End If
    End If
    'Mensual
    If my_tipo_analisis = "Mensual" Then
        ReDim carga_viejo_analisis(0)
        
        mysql = "select month(d.fecha) as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "clientes t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.TIPOCLIE='C'" & Chr$(10)
        'mysql = mysql & "and d.codigo = t.codigo" & Chr$(10)
        
        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and t.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by month(d.fecha),d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
       If mytable.Fields("mes") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
       End If
       If mytable.Fields("moneda") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
       End If
       If mytable.Fields("xtotal") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
       End If
       k = k + 1
       mytable.MoveNext
      Loop

      mytable.Close
     End If
    End If
    'semanal
    If my_tipo_analisis = "Semanal" Then
        ReDim carga_viejo_analisis(0)
        
        mysql = "select DATENAME(weekday,d.fecha) as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "clientes t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.TIPOCLIE='C'" & Chr$(10)
        'mysql = mysql & "and d.codigo = t.codigo" & Chr$(10)
        
        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and t.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by DATENAME(weekday,d.fecha),d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
       If mytable.Fields("mes") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
       End If
       If mytable.Fields("moneda") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
       End If
       If mytable.Fields("xtotal") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
       End If
       k = k + 1
       mytable.MoveNext
      Loop

      mytable.Close

     End If
    End If
   'Diario
    If my_tipo_analisis = "Diario" Then
        ReDim carga_viejo_analisis(0)
        
        mysql = "select day(d.fecha) as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "clientes t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.TIPOCLIE='C'" & Chr$(10)
        'mysql = mysql & "and d.codigo = t.codigo" & Chr$(10)
        
        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and t.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by day(d.fecha),d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
       If mytable.Fields("mes") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
       End If
       If mytable.Fields("moneda") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
       End If
       If mytable.Fields("xtotal") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
       End If
       k = k + 1
       mytable.MoveNext
      Loop

      mytable.Close
   
     End If
    End If
    'Horario
    If my_tipo_analisis = "Horario" Then
        ReDim carga_viejo_analisis(0)
        
        mysql = "select left(d.hora,2) as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "clientes t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.TIPOCLIE='C'" & Chr$(10)
        'mysql = mysql & "and d.codigo = t.codigo" & Chr$(10)

        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and t.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by left(d.hora,2),d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
       If mytable.Fields("mes") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
       End If
       If mytable.Fields("moneda") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
       Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
       End If
       If mytable.Fields("xtotal") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
       Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
       End If
       k = k + 1
       mytable.MoveNext
      Loop

      mytable.Close
      Exit Sub
     End If
    End If
 Exit Sub
End Sub
Public Sub por_vendedor(carga_viejo_analisis() As struct_viejo_analisis, my_tipo_analisis As String, my_moneda As String, _
           my_caja As String, my_servicio As String, my_vendedor As String, my_turno As String, _
           my_familia As String, my_fechai As String, my_fechaf As String, my_tipo As String, _
           my_acu As String, my_codigo As String, salida As Boolean, k As Integer, _
           carga_moneda() As struct_moneda)
           
Dim mysql                     As String
Dim mytable                   As New ADODB.Recordset
Dim mytablex                  As New ADODB.Recordset

   swbuf = "sum(total) as xtotal"
   If cantidad = "Cantidad" Then
   swbuf = "count(numero) as xtotal"
   End If
    'aqui debe cuadrar la tabla tipo con la tabla caja
   'If my_busqueda = "Documentos" Then 'aqui usa la tabla tipo
     If my_tipo_analisis = "Caja" Then
     
        ReDim carga_viejo_analisis(0)
        
        mysql = "select d.TIPOCLIE as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "parameca p," & Chr$(10)
        mysql = mysql & "vendedor t," & Chr$(10)
        mysql = mysql & "producto Pr" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        'mysql = mysql & "and d.vendedor = t.codigo" & Chr$(10)
        mysql = mysql & "and d.CAJA = p.caja" & Chr$(10)

        If my_tipo <> "%" Then
         mysql = mysql & "and d.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and d.codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and pr.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and d.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and p.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and d.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and pr.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by d.TIPOCLIE ,d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
       If mytable.Fields("mes") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
       End If
       If mytable.Fields("moneda") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
       End If
       If mytable.Fields("xtotal") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
       End If
       k = k + 1
       mytable.MoveNext
      Loop
      mytable.Close

     End If
  
    End If
   '**aqui es el tipo
   If my_tipo_analisis = "Tipo" Then
     
        ReDim carga_viejo_analisis(0)
        
        mysql = "select d.tipo as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from  factura d," & Chr$(10)
        mysql = mysql & "Vendedor t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        'mysql = mysql & "and d.vendedor = t.codigo" & Chr$(10)

        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and t.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by d.tipo ,d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
       If mytable.Fields("mes") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
       Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
       End If
       If mytable.Fields("moneda") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
       End If
       If mytable.Fields("xtotal") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
       End If
       k = k + 1
       mytable.MoveNext
      Loop
      
      mytable.Close
     End If
  
    End If
   '**aqui es el usuario
   If my_tipo_analisis = "Usuario" Then
     
        ReDim carga_viejo_analisis(0)
        
        mysql = "select d.usuario as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "Vendedor t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        'mysql = mysql & "and d.vendedor = t.codigo" & Chr$(10)
        
        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and t.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by d.usuario ,d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
        If mytable.Fields("mes") <> "" Then
       carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
       End If
       If mytable.Fields("moneda") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
        Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
       End If
       If mytable.Fields("xtotal") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
       End If
       k = k + 1
       mytable.MoveNext
      Loop

      mytable.Close
     End If
    End If
    '**aqui es para la bodega
    If my_tipo_analisis = "Bodega" Then
     
        ReDim carga_viejo_analisis(0)
        
        mysql = "select d.bodega as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "Vendedor t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        'mysql = mysql & "and d.vendedor = t.codigo" & Chr$(10)
       
        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and t.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by d.bodega ,d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
       If mytable.Fields("mes") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
       End If
       If mytable.Fields("moneda") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
       End If
       If mytable.Fields("xtotal") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
       End If
       k = k + 1
       mytable.MoveNext
      Loop

      mytable.Close
     End If
    End If
    
    '**aqui es para la codigo
    If my_tipo_analisis = "Codigo" Then
     
        ReDim carga_viejo_analisis(0)
        
        mysql = "select d.codigo as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "Vendedor t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        'mysql = mysql & "and d.vendedor = t.codigo" & Chr$(10)
        
        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and t.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by d.codigo,d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
        If mytable.Fields("mes") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
        Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
        End If
        If mytable.Fields("moneda") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
        Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
        End If
        If mytable.Fields("xtotal") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
        Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
        End If
       k = k + 1
       mytable.MoveNext
      Loop

      mytable.Close
     End If
    End If
     'aqui es para el vendedor
     If my_tipo_analisis = "Vendedor" Then
     
        ReDim carga_viejo_analisis(0)
        
        mysql = "select d.vendedor as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "Vendedor t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        'mysql = mysql & "and d.vendedor = t.codigo" & Chr$(10)
        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and t.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by d.vendedor,d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
        If mytable.Fields("mes") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
        Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
        End If
        If mytable.Fields("moneda") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
        Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
        End If
        If mytable.Fields("xtotal") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
        Else
          carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
        End If
       k = k + 1
       mytable.MoveNext
      Loop

      mytable.Close
     End If
    End If
    'aqui es para anual,mensual,semanal,diario,horario
     If my_tipo_analisis = "Anual" Then
     
        ReDim carga_viejo_analisis(0)
        
        mysql = "select year(d.fecha) as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "Vendedor t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        'mysql = mysql & "and d.vendedor = t.codigo" & Chr$(10)
        
        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and t.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by year(d.fecha),d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
       If mytable.Fields("mes") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
       End If
       If mytable.Fields("moneda") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
       End If
       If mytable.Fields("xtotal") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
       End If
       k = k + 1
       mytable.MoveNext
      Loop

      mytable.Close
     End If
    End If
    'mensual
     If my_tipo_analisis = "Mensual" Then
     
        ReDim carga_viejo_analisis(0)
        
        mysql = "select month(d.fecha) as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "Vendedor t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        'mysql = mysql & "and d.vendedor = t.codigo" & Chr$(10)
        
        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and t.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by month(d.fecha),d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
       If mytable.Fields("mes") <> "" Then
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
       Else
         carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
       End If
       If mytable.Fields("moneda") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
       End If
       If mytable.Fields("xtotal") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
       End If
       k = k + 1
       mytable.MoveNext
      Loop

      mytable.Close
     End If
    End If
    'Semanal
    If my_tipo_analisis = "Semanal" Then
     
        ReDim carga_viejo_analisis(0)
        
        mysql = "select DATENAME(weekday,d.fecha) as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "Vendedor t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        'mysql = mysql & "and d.vendedor = t.codigo" & Chr$(10)
        
        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and t.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by DATENAME(weekday,d.fecha),d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
       If mytable.Fields("mes") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
       End If
       If mytable.Fields("moneda") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
       End If
       If mytable.Fields("xtotal") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
       End If
       k = k + 1
       mytable.MoveNext
      Loop
      mytable.Close
      Exit Sub
     End If
    End If
  'Diario
  If my_tipo_analisis = "Diario" Then
     
        ReDim carga_viejo_analisis(0)
        
        mysql = "select day(d.fecha) as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "Vendedor t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        'mysql = mysql & "and d.vendedor = t.codigo" & Chr$(10)
        
        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and t.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by day(d.fecha),d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
       If mytable.Fields("mes") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
       End If
       If mytable.Fields("moneda") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
       End If
       If mytable.Fields("xtotal") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
       End If
       k = k + 1
       mytable.MoveNext
      Loop
      mytable.Close
      Exit Sub
     End If
    End If
    'Horario
    If my_tipo_analisis = "Horario" Then
     
        ReDim carga_viejo_analisis(0)
        
        mysql = "select left(d.hora,2) as mes,d.moneda," & swbuf & Chr$(10)
        mysql = mysql & "from factura d," & Chr$(10)
        mysql = mysql & "Vendedor t" & Chr$(10)
        mysql = mysql & "where d.moneda='" & Format(my_moneda) & "'" & Chr$(10)
        mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "'" & Chr$(10)
        'mysql = mysql & "and d.vendedor = t.codigo" & Chr$(10)
        
        If my_tipo <> "%" Then
         mysql = mysql & "and t.tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)
       End If
       If my_codigo <> "%" Then
          mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)
       End If
       If my_servicio <> "%" Then
         mysql = mysql & "and t.servicio like '" & my_servicio & "'" & Chr$(10)
       End If
       If my_moneda <> "%" Then
        mysql = mysql & "and t.moneda like '" & my_moneda & "'" & Chr$(10)
       End If
       If my_caja <> "%" Then
        mysql = mysql & "and t.caja like '" & my_caja & "'" & Chr$(10)
       End If
      If my_turno <> "%" Then
        mysql = mysql & "and t.turno like '" & my_turno & "'" & Chr$(10)
      End If
      If my_vendedor <> "%" Then
       mysql = mysql & "and d.vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)
      End If
      If my_familia <> "%" Then
        mysql = mysql & "and t.familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)
      End If
      If acu = "Ventas" Then
       mysql = mysql & "and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)
      End If
      If acu = "Compras" Then
        mysql = mysql & "and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)
      End If
    
      mysql = mysql & "group by left(d.hora,2),d.moneda" & Chr$(10)

     mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
     If mytable.EOF Then
        salida = False
        Exit Sub
     Else
       salida = True
       mytable.MoveFirst
       k = 0
      Do Until mytable.EOF
        If k > 0 Then
          ReDim Preserve carga_viejo_analisis(UBound(carga_viejo_analisis) + 1)
        End If
       If mytable.Fields("mes") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = mytable.Fields("mes")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).mes = ""
       End If
       If mytable.Fields("moneda") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = mytable.Fields("moneda")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).moneda = ""
       End If
       If mytable.Fields("xtotal") <> "" Then
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = mytable.Fields("xtotal")
       Else
        carga_viejo_analisis(UBound(carga_viejo_analisis)).total = 0
       End If
       k = k + 1
       mytable.MoveNext
      Loop
      mytable.Close
     End If
    End If
    
Exit Sub

End Sub
Public Sub detalle_chart_sele2(objWorkBook As Excel.Workbook, carga_viejo_analisis() As struct_viejo_analisis, _
            k As Integer, v As Integer)

                             
Dim mysql                     As String
Dim mytable                   As New ADODB.Recordset

On Error GoTo detalle_chart_sele2


v = 13
h = 0
        For j = 0 To k - 1
         objWorksheet.Cells(v, h + 1) = carga_viejo_analisis(j).mes
         If carga_viejo_analisis(j).moneda = "S" Then
           objWorksheet.Cells(v, h + 2) = "Soles"
         Else
          objWorksheet.Cells(v, h + 2) = "Dolares"
         End If
         objWorksheet.Cells(v, h + 3) = carga_viejo_analisis(j).total
         
         v = v + 1
        Next j


Exit Sub
detalle_chart_sele2:
Select Case Err.Number
        Case 1004
            Resume Next
        Case Else
            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select
Exit Sub
End Sub
Public Sub crear_chart2_soles(objWorkBook As Excel.Workbook, v As Integer, h As Integer, _
           my_tipo_analisis As String)
                        
Dim objExcel                           As Excel.Application
Dim objWorksheet                       As Excel.Worksheet
Dim chtobj                             As ChartObject
Dim my_final                           As Integer

 On Error GoTo crear_chart

' ir = objWorksheet.Cells(v, 1).End(xlUp).Row
' Set chtRng = Range("A13:C" & ir)

 
 my_final = v - 1
 
 Set chtRng = Range("A13:C" & my_final)
 
 objWorkBook.Charts.Add
 objWorkBook.ActiveChart.SetSourceData Source:=objWorkBook.Sheets("Hoja1").Range("A13:C" & my_final), PlotBy:=xlColumns
 objWorkBook.ActiveChart.SeriesCollection(1).Name = "=""Fecha"""


 Charts(1).Activate
 With ActiveChart
  .HasTitle = True
  .ChartStyle = 34
  .ChartTitle.Characters.Text = "Suma Total Factura En Soles"
  .Axes(xlCategory, xlPrimary).HasTitle = True
  .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = my_tipo_analisis
  '"Mes"
  .Axes(xlValue, xlPrimary).HasTitle = True
  .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Total"

 End With
 Exit Sub
crear_chart:
 Select Case Err.Number
        Case 1004
            Resume Next
        Case Else
           ' MsgBox Err.Number & vbcrlf & Err.Description
    End Select
 Exit Sub
End Sub
Public Sub crear_chart2_dolar(objWorkBook As Excel.Workbook, v As Integer, h As Integer, _
           my_tipo_analisis As String)
                        
Dim objExcel                           As Excel.Application
Dim objWorksheet                       As Excel.Worksheet
Dim chtobj                             As ChartObject
Dim my_final                           As Integer

 On Error GoTo crear_chart

 ir = objWorksheet.Cells(v, 1).End(xlUp).Row
 Set chtRng = Range("A13:C" & ir)

 my_final = v - 1
 
 objWorkBook.Charts.Add
 objWorkBook.ActiveChart.SetSourceData Source:=objWorkBook.Sheets("Hoja1").Range("A13:C" & my_final), PlotBy:=xlColumns
 objWorkBook.ActiveChart.SeriesCollection(1).Name = "=""Fecha"""


 With ActiveChart
  .HasTitle = True
  .ChartStyle = 34
  .ChartTitle.Characters.Text = "Suma Total Factura En Dolares"
  .Axes(xlCategory, xlPrimary).HasTitle = True
  .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Mes"
  .Axes(xlValue, xlPrimary).HasTitle = True
  .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Total"

 End With
crear_chart:
 Select Case Err.Number
        Case 1004
            Resume Next
        Case Else
           ' MsgBox Err.Number & vbcrlf & Err.Description
    End Select
End Sub
Public Sub cargar_tgrafico(my_struct_reportG() As struct_reportG, _
            k As Integer, v As Integer)

                             
Dim mysql                     As String
Dim mytable                   As New ADODB.Recordset

On Error GoTo detalle_chart_sele2


v = 11
h = 0
        For j = 0 To k - 1
         objWorksheet.Cells(v, h + 1) = my_struct_reportG(j).fecha
         objWorksheet.Cells(v, h + 2) = my_struct_reportG(j).cajaVentas
         objWorksheet.Cells(v, h + 3) = my_struct_reportG(j).TotalVenta
         objWorksheet.Cells(v, h + 4) = my_struct_reportG(j).cajaCompras
         objWorksheet.Cells(v, h + 3) = my_struct_reportG(j).TotalCompras
         v = v + 1
        Next j


Exit Sub
detalle_chart_sele2:
Select Case Err.Number
        Case 1004
            Resume Next
        Case Else
            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select
Exit Sub
End Sub
