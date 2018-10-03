Attribute VB_Name = "Mtfactura"

'inicio 25/07/2017 pll
'para la grid consulta
Const N_col_nombre = 0

Const N_col_ruc = 1

Const N_col_codigo1 = 2

Const N_col_fpago = 3

Type struc_busca_cliente

    nombre                        As String
    RUC                           As String
    codigo1                       As String
    fpago                         As String
    tipo                          As String * 1
    direccion                     As String
    correo                        As String
    dpto                          As String
    distrito                      As String
    provincia                     As String

End Type

Global my_carga_busca_cliente() As struc_busca_cliente

'inicio 26/07/2017 pll esto es para los productos
'para la grid Productos
Const N_cols_descripcion = 0

Const N_cols_codProducto = 1

Const N_cols_marca = 2

Const N_cols_unidad = 3

Const N_cols_factor = 4

Const N_cols_precio = 5

Const N_cols_moneda = 6

Const N_cols_familia = 7

Const N_cols_subfamilia = 8

Const N_cols_barra = 9

Const N_cols_igv = 10

Const N_cols_local = 11

Type struc_busca_producto

    descripcion                       As String
    codProducto                       As String
    marca                             As String
    unidad                            As String
    factor                            As String
    precio                            As String
    moneda                            As String
    familia                           As String
    subfamilia                        As String
    barra                             As String
    igv                               As Double
    local                             As String

End Type

Global my_carga_busca_producto() As struc_busca_producto

'fin 26/07/2017 pll
'inicio 27/07/2017 pll
'esto Exit Sub para producto vs.precio
Const N_columna_codigo = 0

Const N_columna_unidad1 = 1

Const N_columna_pventa1 = 2

Const N_columna_unidad2 = 3

Const N_columna_pventa2 = 4

Const N_columna_unidad3 = 5

Const N_columna_pventa3 = 6

Const N_columna_unidad4 = 7

Const N_columna_pventa4 = 8

Const N_columna_unidad5 = 9

Const N_columna_pventa5 = 10

Const N_columna_unidad6 = 11

Const N_columna_pventa6 = 12

Const N_columna_unidad7 = 13

Const N_columna_pventa7 = 14

Const N_columna_unidad8 = 15

Const N_columna_pventa8 = 16

Const N_columna_unidad9 = 17

Const N_columna_pventa9 = 18

Const N_columna_unidad10 = 19

Const N_columna_pventa10 = 20

Type struc_producto_precio

    codigo                            As String
    unidad1                           As String
    pventa1                           As Double
    unidad2                           As String
    pventa2                           As Double
    unidad3                           As String
    pventa3                           As Double
    unidad4                           As String
    pventa4                           As Double
    unidad5                           As String
    pventa5                           As Double
    unidad6                           As String
    pventa6                           As Double
    unidad7                           As String
    pventa7                           As Double
    unidad8                           As String
    pventa8                           As Double
    unidad9                           As String
    pventa9                           As Double
    unidad10                          As String
    pventa10                          As Double

End Type

Global my_carga_producto_precio() As struc_producto_precio

'fin 27/07/2017 pll
'inicio 01/08/2017 pll para la ventas
Type struc_tipoDocumento

    descripcion                       As String
    tipo                              As String

End Type

Global my_carga_tipoDocumento() As struc_tipoDocumento

'fin 01/08/2017 pll para la ventas
'inicio 01/08/2017 pll proveedor y vendedor
Type struc_proveedor

    nombre                              As String
    codigo                              As String

End Type

Global my_carga_struc_proveedor() As struc_proveedor

'grid proveedor
Const N_colm_descripcion = 0

Const N_colm_codProducto = 1

'fin 01/08/2017 pll
'inicio 17/08/2017 pll
'para la grid trasportista
Const N_c_nombre = 0

Const N_c_codigo = 1

Type struc_transporte

    nombre                              As String
    codigo                              As String

End Type

Global my_struc_transporte() As struc_transporte

'Public my_bodegaf                   As String
'fin 17/08/2017 pll
Type struc_vendedor

    codigo                            As String
    nombre                            As String

End Type

Global my_struc_vendedor() As struc_vendedor

Global array_vendedor      As Integer

Global vendedor            As String

'inicio 28/11/2017 pll
Global my_vdolar           As String

'fin 28/11/2017 pll
Public Function ini_grid_bus_cliente(my_grid As MSFlexGrid)

    'utilizo per eliminare_fattura<--voce
    my_grid.Clear
    'inizializzazione Grid
    my_grid.rows = 2
    my_grid.FixedRows = 1
    my_grid.Cols = 4 'aqui se aumenta
    my_grid.FixedCols = 0

    my_grid.Row = 0
    my_grid.Col = N_col_nombre 'dflag
    my_grid.Text = "Nombre"
    my_grid.ColWidth(my_grid.Col) = 4800
    my_grid.FixedAlignment(my_grid.Col) = 0
    my_grid.ColAlignment(my_grid.Col) = 0

    my_grid.Row = 0
    my_grid.Col = N_col_ruc 'yausado
    my_grid.Text = "Ruc"
    my_grid.ColWidth(my_grid.Col) = 3500
    my_grid.FixedAlignment(my_grid.Col) = 0
    my_grid.ColAlignment(my_grid.Col) = 0

    my_grid.Row = 0
    my_grid.Col = N_col_codigo1 'estado 1
    my_grid.Text = "Codigo"
    my_grid.ColWidth(my_grid.Col) = 2500
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_col_fpago 'estado 1
    my_grid.Text = "Fpago"
    my_grid.ColWidth(my_grid.Col) = 0
    my_grid.FixedAlignment(my_grid.Col) = 0
    my_grid.ColAlignment(my_grid.Col) = 0

    Exit Function

End Function

Public Function carica_busca_cliente(my_grid As MSFlexGrid, _
                                     my_carga_busca_cliente() As struc_busca_cliente, _
                                     k As Integer)

    For I = 0 To k - 1
        'c = c + 1

        my_grid.AddItem ""
        my_grid.Row = my_grid.rows - 1

        my_grid.Col = N_col_nombre
        my_grid.Text = my_carga_busca_cliente(I).nombre

        my_grid.Col = N_col_ruc
        my_grid.Text = my_carga_busca_cliente(I).RUC

        my_grid.Col = N_col_codigo1
        my_grid.Text = my_carga_busca_cliente(I).codigo1

        my_grid.Col = N_col_fpago
        my_grid.Text = my_carga_busca_cliente(I).fpago
    Next I

    my_grid.Row = 1
    my_grid.Col = 0

    If my_grid.Text = "" Then
        my_grid.RemoveItem 1

    End If

    Exit Function

End Function

Public Function busca_cliente(my_busqueda As String, _
                              my_carga_busca_cliente() As struc_busca_cliente, _
                              salida As Boolean, _
                              k As Integer)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo busca_cliente_err

    ReDim my_carga_busca_cliente(0)

    mysql = "select Nombre,Codigo,Codigo1," & Chr$(10)
    mysql = mysql & "fpago," & Chr$(10)
    mysql = mysql & "isnull(tipo,'vacio') as tipo," & Chr$(10) 'número correlativo item1"
    mysql = mysql & "dpto , distrito, " & Chr$(10)
    mysql = mysql & "direccion,correo,provincia" & Chr$(10)
    mysql = mysql & "from clientes" & Chr$(10)
    mysql = mysql & "where Nombre" & " like '" & my_busqueda & "'" & Chr$(10)
    mysql = mysql & "OR Codigo" & " like '" & my_busqueda & "'" & Chr$(10)

    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF Then
        salida = False
        Exit Function
    Else
        salida = True
        mytablex.MoveFirst
        k = 0

        Do Until mytablex.EOF

            'para cargar la msgrid
            If k > 0 Then
                ReDim Preserve my_carga_busca_cliente(UBound(my_carga_busca_cliente) + 1)

            End If

            If mytablex.Fields("Nombre") <> "" Then
                my_carga_busca_cliente(UBound(my_carga_busca_cliente)).nombre = mytablex.Fields("Nombre")
            Else
                my_carga_busca_cliente(UBound(my_carga_busca_cliente)).nombre = ""

            End If
      
            'factura exportacion
            If mytablex.Fields("Codigo") <> "" Then ' 1006899020
                my_carga_busca_cliente(UBound(my_carga_busca_cliente)).RUC = my_busqueda
            Else
                my_carga_busca_cliente(UBound(my_carga_busca_cliente)).RUC = ""

            End If

            If mytablex.Fields("Codigo1") <> "" Then
                my_carga_busca_cliente(UBound(my_carga_busca_cliente)).codigo1 = mytablex.Fields("Codigo1")
            Else
                my_carga_busca_cliente(UBound(my_carga_busca_cliente)).codigo1 = ""

            End If

            If mytablex.Fields("fpago") <> "" Then
                my_carga_busca_cliente(UBound(my_carga_busca_cliente)).fpago = mytablex.Fields("fpago")
            Else
                my_carga_busca_cliente(UBound(my_carga_busca_cliente)).fpago = 1

            End If

            If mytablex.Fields("tipo") <> "" Then
                my_carga_busca_cliente(UBound(my_carga_busca_cliente)).tipo = mytablex.Fields("tipo")
            Else
                my_carga_busca_cliente(UBound(my_carga_busca_cliente)).tipo = ""

            End If

            If mytablex.Fields("direccion") <> "" Then
                my_carga_busca_cliente(UBound(my_carga_busca_cliente)).direccion = mytablex.Fields("direccion")
            Else
                my_carga_busca_cliente(UBound(my_carga_busca_cliente)).direccion = ""

            End If

            If mytablex.Fields("correo") <> "" Then
                my_carga_busca_cliente(UBound(my_carga_busca_cliente)).correo = mytablex.Fields("correo")
            Else
                my_carga_busca_cliente(UBound(my_carga_busca_cliente)).correo = ""

            End If

            If mytablex.Fields("dpto") <> "" Then
                my_carga_busca_cliente(UBound(my_carga_busca_cliente)).dpto = mytablex.Fields("dpto")
            Else
                my_carga_busca_cliente(UBound(my_carga_busca_cliente)).dpto = ""

            End If

            If mytablex.Fields("distrito") <> "" Then
                my_carga_busca_cliente(UBound(my_carga_busca_cliente)).distrito = mytablex.Fields("distrito")
            Else
                my_carga_busca_cliente(UBound(my_carga_busca_cliente)).distrito = ""

            End If
  
            If mytablex.Fields("provincia") <> "" Then
                my_carga_busca_cliente(UBound(my_carga_busca_cliente)).provincia = mytablex.Fields("provincia")
            Else
                my_carga_busca_cliente(UBound(my_carga_busca_cliente)).provincia = ""

            End If
  
            k = k + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

    Exit Function

busca_cliente_err:
    MsgBox "aviso en sql_cabeza   " & error$, 48, "Aviso"
    Exit Function

End Function

'inizializa para el producto 26/07/2017 pll
Public Function ini_grid_bus_producto(my_grid As MSFlexGrid)

    'utilizo per eliminare_fattura<--voce
    my_grid.Clear
    'inizializzazione Grid
    my_grid.rows = 2
    my_grid.FixedRows = 1
    my_grid.Cols = 13 'aqui se aumenta
    my_grid.FixedCols = 0

    my_grid.Row = 0
    my_grid.Col = N_cols_descripcion 'descripcion
    my_grid.Text = "Descripcion"
    my_grid.ColWidth(my_grid.Col) = 3200
    my_grid.FixedAlignment(my_grid.Col) = 0
    my_grid.ColAlignment(my_grid.Col) = 0

    my_grid.Row = 0
    my_grid.Col = N_cols_codProducto 'codProducto
    my_grid.Text = "Cod.Producto"
    my_grid.ColWidth(my_grid.Col) = 1200
    my_grid.FixedAlignment(my_grid.Col) = 0
    my_grid.ColAlignment(my_grid.Col) = 0

    my_grid.Row = 0
    my_grid.Col = N_cols_marca 'marca
    my_grid.Text = "Marca"
    my_grid.ColWidth(my_grid.Col) = 1200
    my_grid.FixedAlignment(my_grid.Col) = 2
    my_grid.ColAlignment(my_grid.Col) = 2

    my_grid.Row = 0
    my_grid.Col = N_cols_unidad 'unidad
    my_grid.Text = "Unidad"
    my_grid.ColWidth(my_grid.Col) = 800
    my_grid.FixedAlignment(my_grid.Col) = 0
    my_grid.ColAlignment(my_grid.Col) = 0

    my_grid.Row = 0
    my_grid.Col = N_cols_factor 'factor
    my_grid.Text = "Factor"
    my_grid.ColWidth(my_grid.Col) = 800
    my_grid.FixedAlignment(my_grid.Col) = 0
    my_grid.ColAlignment(my_grid.Col) = 0

    my_grid.Row = 0
    my_grid.Col = N_cols_precio 'precio
    my_grid.Text = "Precio"
    my_grid.ColWidth(my_grid.Col) = 1800
    my_grid.FixedAlignment(my_grid.Col) = 0
    my_grid.ColAlignment(my_grid.Col) = 0

    my_grid.Row = 0
    my_grid.Col = N_cols_moneda 'moneda
    my_grid.Text = "Moneda"
    my_grid.ColWidth(my_grid.Col) = 900
    my_grid.FixedAlignment(my_grid.Col) = 0
    my_grid.ColAlignment(my_grid.Col) = 0

    my_grid.Row = 0
    my_grid.Col = N_cols_familia 'familia
    my_grid.Text = "Familia"
    my_grid.ColWidth(my_grid.Col) = 1800
    my_grid.FixedAlignment(my_grid.Col) = 0
    my_grid.ColAlignment(my_grid.Col) = 0

    my_grid.Row = 0
    my_grid.Col = N_cols_subfamilia 'familia
    my_grid.Text = "SubFamilia"
    my_grid.ColWidth(my_grid.Col) = 1800
    my_grid.FixedAlignment(my_grid.Col) = 0
    my_grid.ColAlignment(my_grid.Col) = 0

    my_grid.Row = 0
    my_grid.Col = N_cols_barra 'familia
    my_grid.Text = "Barra"
    my_grid.ColWidth(my_grid.Col) = 1800
    my_grid.FixedAlignment(my_grid.Col) = 0
    my_grid.ColAlignment(my_grid.Col) = 0

    my_grid.Row = 0
    my_grid.Col = N_cols_igv 'familia
    my_grid.Text = "IGV"
    my_grid.ColWidth(my_grid.Col) = 1800
    my_grid.FixedAlignment(my_grid.Col) = 0
    my_grid.ColAlignment(my_grid.Col) = 0

    my_grid.Row = 0
    my_grid.Col = N_cols_local 'local
    my_grid.Text = "Local"
    my_grid.ColWidth(my_grid.Col) = 0
    my_grid.FixedAlignment(my_grid.Col) = 0
    my_grid.ColAlignment(my_grid.Col) = 0

    Exit Function

End Function

Sub ver_presenta()

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    buf = "" & DBGrid2.columns("producto")
    presenta = ""
    precio = ""
    mytablex.Open "SELECT * FROM producto where  producto='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        presenta = "" & mytablex.Fields("presenta")

    End If

    mytablex.Close
    mytablex.Open "SELECT * FROM precios where  producto='" & buf & "' and local='" & local1 & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        precio = "" & mytablex.Fields("pventa1")

    End If

    mytablex.Close

End Sub

Public Sub buscar_producto(my_busqueda As String, _
                           my_carga_busca_producto() As struc_busca_producto, _
                           salida As Boolean, _
                           k As Integer)
           
    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    ReDim my_carga_busca_producto(0)

    mysql = "SELECT a.descripcio,a.producto,a.marca,a.unidad,a.factor," & Chr$(10)
    mysql = mysql & "a.monedac ,a.familia,a.subfamilia, " & Chr$(10)
    mysql = mysql & "a.barras,a.igv,B.PVENTA1,b.local" & Chr$(10)
    mysql = mysql & "FROM" & Chr$(10)
    mysql = mysql & "(SELECT descripcio,producto,marca,unidad,factor," & Chr$(10)
    mysql = mysql & " monedac , familia, subfamilia, Barras, igv" & Chr$(10)
    mysql = mysql & "From producto" & Chr$(10)
    mysql = mysql & "where descripcio" & " like '" & my_busqueda & "'" & Chr$(10)
    mysql = mysql & "or producto" & " like  '" & my_busqueda & "'" & ")a," & Chr$(10)
    mysql = mysql & "(select PVENTA1,PRODUCTO,local" & Chr$(10)
    mysql = mysql & " from precios p)b" & Chr$(10)
    mysql = mysql & "Where a.producto = b.producto" & Chr$(10)

    mysql = mysql & "order by a.descripcio asc" & Chr$(10)

    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF Then
        salida = False
        Exit Sub
    Else
        salida = True
        mytablex.MoveFirst
        k = 0

        Do Until mytablex.EOF

            'para cargar la msgrid
            If k > 0 Then
                ReDim Preserve my_carga_busca_producto(UBound(my_carga_busca_producto) + 1)

            End If

            If mytablex.Fields("descripcio") <> "" Then
                my_carga_busca_producto(UBound(my_carga_busca_producto)).descripcion = mytablex.Fields("descripcio")
            Else
                my_carga_busca_producto(UBound(my_carga_busca_producto)).descripcion = ""

            End If
      
            If mytablex.Fields("producto") <> "" Then
                my_carga_busca_producto(UBound(my_carga_busca_producto)).codProducto = mytablex.Fields("producto")
            Else
                my_carga_busca_producto(UBound(my_carga_busca_producto)).codProducto = ""

            End If
      
            If mytablex.Fields("marca") <> "" Then
                my_carga_busca_producto(UBound(my_carga_busca_producto)).marca = mytablex.Fields("marca")
            Else
                my_carga_busca_producto(UBound(my_carga_busca_producto)).marca = ""

            End If
      
            If mytablex.Fields("unidad") <> "" Then
                my_carga_busca_producto(UBound(my_carga_busca_producto)).unidad = mytablex.Fields("unidad")
            Else
                my_carga_busca_producto(UBound(my_carga_busca_producto)).unidad = ""

            End If
      
            If mytablex.Fields("factor") <> "" Then
                my_carga_busca_producto(UBound(my_carga_busca_producto)).factor = mytablex.Fields("factor")
            Else
                my_carga_busca_producto(UBound(my_carga_busca_producto)).factor = ""

            End If
      
            If mytablex.Fields("PVENTA1") <> "" Then
                my_carga_busca_producto(UBound(my_carga_busca_producto)).precio = mytablex.Fields("PVENTA1")
            Else
                my_carga_busca_producto(UBound(my_carga_busca_producto)).precio = ""

            End If
      
            If mytablex.Fields("monedac") <> "" Then
                my_carga_busca_producto(UBound(my_carga_busca_producto)).moneda = mytablex.Fields("monedac")
            Else
                my_carga_busca_producto(UBound(my_carga_busca_producto)).moneda = ""

            End If
      
            If mytablex.Fields("familia") <> "" Then
                my_carga_busca_producto(UBound(my_carga_busca_producto)).familia = mytablex.Fields("familia")
            Else
                my_carga_busca_producto(UBound(my_carga_busca_producto)).familia = ""

            End If
      
            If mytablex.Fields("subfamilia") <> "" Then
                my_carga_busca_producto(UBound(my_carga_busca_producto)).subfamilia = mytablex.Fields("subfamilia")
            Else
                my_carga_busca_producto(UBound(my_carga_busca_producto)).subfamilia = ""

            End If
         
            If mytablex.Fields("barras") <> "" Then
                my_carga_busca_producto(UBound(my_carga_busca_producto)).barra = mytablex.Fields("barras")
            Else
                my_carga_busca_producto(UBound(my_carga_busca_producto)).barra = ""

            End If
      
            If mytablex.Fields("igv") <> "" Then
                my_carga_busca_producto(UBound(my_carga_busca_producto)).igv = mytablex.Fields("igv")
            Else
                my_carga_busca_producto(UBound(my_carga_busca_producto)).igv = 0

            End If

            If mytablex.Fields("local") <> "" Then
                my_carga_busca_producto(UBound(my_carga_busca_producto)).local = mytablex.Fields("local")
            Else
                my_carga_busca_producto(UBound(my_carga_busca_producto)).local = ""

            End If
      
            k = k + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

    Exit Sub

End Sub

Public Function carica_busca_producto(my_grid As MSFlexGrid, _
                                      my_carga_busca_producto() As struc_busca_producto, _
                                      k As Integer)

    For I = 0 To k - 1
        'c = c + 1

        my_grid.AddItem ""
        my_grid.Row = my_grid.rows - 1

        my_grid.Col = N_cols_descripcion 'descripcion
        my_grid.Text = my_carga_busca_producto(I).descripcion

        my_grid.Col = N_cols_codProducto 'codProducto
        my_grid.Text = my_carga_busca_producto(I).codProducto

        my_grid.Col = N_cols_marca 'marca
        my_grid.Text = my_carga_busca_producto(I).marca

        my_grid.Col = N_cols_unidad 'unidad
        my_grid.Text = my_carga_busca_producto(I).unidad

        my_grid.Col = N_cols_factor 'factor
        my_grid.Text = my_carga_busca_producto(I).factor

        my_grid.Col = N_cols_precio 'precio
        my_grid.Text = my_carga_busca_producto(I).precio

        my_grid.Col = N_cols_moneda 'moneda
        my_grid.Text = my_carga_busca_producto(I).moneda

        my_grid.Col = N_cols_familia 'familia
        my_grid.Text = my_carga_busca_producto(I).familia

        my_grid.Col = N_cols_subfamilia 'familia
        my_grid.Text = my_carga_busca_producto(I).subfamilia

        my_grid.Col = N_cols_barra 'familia
        my_grid.Text = my_carga_busca_producto(I).barra

        my_grid.Col = N_cols_igv 'familia
        my_grid.Text = my_carga_busca_producto(I).igv

        my_grid.Col = N_cols_local 'local
        my_grid.Text = my_carga_busca_producto(I).local

    Next I

    my_grid.Row = 1
    my_grid.Col = 0

    If my_grid.Text = "" Then
        my_grid.RemoveItem 1

    End If

    Exit Function

End Function

Public Function producto_precio(my_codigo As String, _
                                my_carga_producto_precio() As struc_producto_precio, _
                                salida As Boolean, _
                                k As Integer, _
                                my_local As String)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo busca_cliente_err

    ReDim my_carga_producto_precio(0)

    mysql = "SELECT producto,Unidad1,Pventa1,Unidad2,Pventa2,Unidad3,Pventa3,Unidad4,Pventa4," & Chr$(10)
    mysql = mysql & "UNIDAD5,PVENTA5,UNIDAD6,PVENTA6,UNIDAD7,PVENTA7,UNIDAD8,PVENTA8," & Chr$(10)
    mysql = mysql & "UNIDAD9,PVENTA9,UNIDAD10,pventa10" & Chr$(10)
    mysql = mysql & "from Precios " & Chr$(10)
    mysql = mysql & "where producto='" & my_codigo & "'" & Chr$(10)

    If my_local <> "" Then
        mysql = mysql & "and local='" & my_local & "'" & Chr$(10)
    Else
        mysql = mysql & "and local like '%'" & Chr$(10)

    End If

    mysql = mysql & "order by Pventa1 desc" & Chr$(10)

    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF Then
        salida = False
        Exit Function
    Else
        salida = True
        mytablex.MoveFirst
        k = 0

        Do Until mytablex.EOF

            'para cargar la msgrid
            If k > 0 Then
                ReDim Preserve my_carga_producto_precio(UBound(my_carga_producto_precio) + 1)

            End If

            ''
            If mytablex.Fields("producto") <> "" Then
                my_carga_producto_precio(UBound(my_carga_producto_precio)).codigo = mytablex.Fields("producto")
            Else
                my_carga_producto_precio(UBound(my_carga_producto_precio)).codigo = ""

            End If

            ''
            If mytablex.Fields("unidad1") <> "" Then
                my_carga_producto_precio(UBound(my_carga_producto_precio)).unidad1 = mytablex.Fields("unidad1")
            Else
                my_carga_producto_precio(UBound(my_carga_producto_precio)).unidad1 = ""

            End If

            If mytablex.Fields("pventa1") <> "" Then
                my_carga_producto_precio(UBound(my_carga_producto_precio)).pventa1 = mytablex.Fields("pventa1")
            Else
                my_carga_producto_precio(UBound(my_carga_producto_precio)).pventa1 = 0

            End If

            If mytablex.Fields("unidad2") <> "" Then
                my_carga_producto_precio(UBound(my_carga_producto_precio)).unidad2 = mytablex.Fields("unidad2")
            Else
                my_carga_producto_precio(UBound(my_carga_producto_precio)).unidad2 = ""

            End If

            If mytablex.Fields("pventa2") <> "" Then
                my_carga_producto_precio(UBound(my_carga_producto_precio)).pventa2 = mytablex.Fields("pventa2")
            Else
                my_carga_producto_precio(UBound(my_carga_producto_precio)).pventa2 = 0

            End If
      
            If mytablex.Fields("unidad3") <> "" Then
                my_carga_producto_precio(UBound(my_carga_producto_precio)).unidad3 = mytablex.Fields("unidad3")
            Else
                my_carga_producto_precio(UBound(my_carga_producto_precio)).unidad3 = ""

            End If
      
            If mytablex.Fields("pventa3") <> "" Then
                my_carga_producto_precio(UBound(my_carga_producto_precio)).pventa3 = mytablex.Fields("pventa3")
            Else
                my_carga_producto_precio(UBound(my_carga_producto_precio)).pventa3 = 0

            End If
      
            If mytablex.Fields("unidad4") <> "" Then
                my_carga_producto_precio(UBound(my_carga_producto_precio)).unidad4 = mytablex.Fields("unidad4")
            Else
                my_carga_producto_precio(UBound(my_carga_producto_precio)).unidad4 = ""

            End If
    
            If mytablex.Fields("pventa4") <> "" Then
                my_carga_producto_precio(UBound(my_carga_producto_precio)).pventa4 = mytablex.Fields("pventa4")
            Else
                my_carga_producto_precio(UBound(my_carga_producto_precio)).pventa4 = 0

            End If
    
            If mytablex.Fields("unidad5") <> "" Then
                my_carga_producto_precio(UBound(my_carga_producto_precio)).unidad5 = mytablex.Fields("unidad5")
            Else
                my_carga_producto_precio(UBound(my_carga_producto_precio)).unidad5 = ""

            End If
    
            If mytablex.Fields("pventa5") <> "" Then
                my_carga_producto_precio(UBound(my_carga_producto_precio)).pventa5 = mytablex.Fields("pventa5")
            Else
                my_carga_producto_precio(UBound(my_carga_producto_precio)).pventa5 = 0

            End If
    
            If mytablex.Fields("unidad6") <> "" Then
                my_carga_producto_precio(UBound(my_carga_producto_precio)).unidad6 = mytablex.Fields("unidad6")
            Else
                my_carga_producto_precio(UBound(my_carga_producto_precio)).unidad6 = ""

            End If
    
            If mytablex.Fields("pventa6") <> "" Then
                my_carga_producto_precio(UBound(my_carga_producto_precio)).pventa6 = mytablex.Fields("pventa6")
            Else
                my_carga_producto_precio(UBound(my_carga_producto_precio)).pventa6 = 0

            End If
    
            If mytablex.Fields("unidad7") <> "" Then
                my_carga_producto_precio(UBound(my_carga_producto_precio)).unidad7 = mytablex.Fields("unidad7")
            Else
                my_carga_producto_precio(UBound(my_carga_producto_precio)).unidad7 = ""

            End If
    
            If mytablex.Fields("pventa7") <> "" Then
                my_carga_producto_precio(UBound(my_carga_producto_precio)).pventa7 = mytablex.Fields("pventa7")
            Else
                my_carga_producto_precio(UBound(my_carga_producto_precio)).pventa7 = 0

            End If
    
            If mytablex.Fields("unidad8") <> "" Then
                my_carga_producto_precio(UBound(my_carga_producto_precio)).unidad8 = mytablex.Fields("unidad8")
            Else
                my_carga_producto_precio(UBound(my_carga_producto_precio)).unidad8 = ""

            End If
    
            If mytablex.Fields("pventa8") <> "" Then
                my_carga_producto_precio(UBound(my_carga_producto_precio)).pventa8 = mytablex.Fields("pventa8")
            Else
                my_carga_producto_precio(UBound(my_carga_producto_precio)).pventa8 = 0

            End If
    
            If mytablex.Fields("unidad9") <> "" Then
                my_carga_producto_precio(UBound(my_carga_producto_precio)).unidad9 = mytablex.Fields("unidad9")
            Else
                my_carga_producto_precio(UBound(my_carga_producto_precio)).unidad9 = ""

            End If

            If mytablex.Fields("pventa9") <> "" Then
                my_carga_producto_precio(UBound(my_carga_producto_precio)).pventa9 = mytablex.Fields("pventa9")
            Else
                my_carga_producto_precio(UBound(my_carga_producto_precio)).pventa8 = 0

            End If

            If mytablex.Fields("unidad10") <> "" Then
                my_carga_producto_precio(UBound(my_carga_producto_precio)).unidad10 = mytablex.Fields("unidad10")
            Else
                my_carga_producto_precio(UBound(my_carga_producto_precio)).unidad10 = ""

            End If

            If mytablex.Fields("pventa10") <> "" Then
                my_carga_producto_precio(UBound(my_carga_producto_precio)).pventa10 = mytablex.Fields("pventa10")
            Else
                my_carga_producto_precio(UBound(my_carga_producto_precio)).pventa10 = 0

            End If
    
            k = k + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

    Exit Function

busca_cliente_err:
    MsgBox "aviso en sql_cabeza   " & error$, 48, "Aviso"
    Exit Function

End Function

'aqui para buscar tipo documento
Public Function busca_cod_tipoProducto(my_descripcion As String, _
                                       my_tipoDocumento As String, _
                                       salida As Boolean)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo busca_cod_tipoProducto

    ReDim my_carga_producto_precio(0)

    mysql = "SELECT Tipo " & Chr$(10)
    mysql = mysql & "from Tipo" & Chr$(10)
    mysql = mysql & "where Descripcio='" & my_descripcion & "'" & Chr$(10)

    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF Then
        salida = False
        Exit Function
    Else
        salida = True
        mytablex.MoveFirst
        k = 0

        Do Until mytablex.EOF
            my_tipoDocumento = mytablex.Fields("tipo")
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

    Exit Function

busca_cod_tipoProducto:
    MsgBox "aviso en sql_cabeza   " & error$, 48, "Aviso"
    Exit Function

End Function

'carga tipo documento
Public Function carga_TipoD(xbuf As String, _
                            my_carga_tipoDocumento() As struc_tipoDocumento, _
                            salida As Boolean, _
                            k As Integer)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo carga_tipoDocumento

    ReDim my_carga_tipoDocumento(0)

    mysql = "select Descripcio,tipo  from Tipo where " & xbuf & Chr$(10)

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
                ReDim Preserve my_carga_tipoDocumento(UBound(my_carga_tipoDocumento) + 1)

            End If

            If mytablex.Fields("Descripcio") <> "" Then
                my_carga_tipoDocumento(UBound(my_carga_tipoDocumento)).descripcion = mytablex.Fields("Descripcio")
            Else
                my_carga_tipoDocumento(UBound(my_carga_tipoDocumento)).descripcion = ""

            End If
       
            If mytablex.Fields("tipo") <> "" Then
                my_carga_tipoDocumento(UBound(my_carga_tipoDocumento)).tipo = mytablex.Fields("tipo")
            Else
                my_carga_tipoDocumento(UBound(my_carga_tipoDocumento)).tipo = ""

            End If
      
            k = k + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

    Exit Function
carga_tipoDocumento:
    MsgBox "carga_tipoDocumento " & error$, 48, "Aviso"
    Exit Function

End Function

'carga proveedor
Public Function carga_proveedor(my_busqueda As String, _
                                my_carga_struc_proveedor() As struc_proveedor, _
                                salida As Boolean, _
                                k As Integer, _
                                my_report)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo carga_tipoDocumento

    ReDim my_carga_struc_proveedor(0)

    'mysql = "select Nombre,codigo from PROVEEDO" & Chr$(10) 'cgusuario
    mysql = "select Nombre,codigo from " & my_report & Chr$(10)
    'mysql = "select Nombre,codigo from VENDEDOR" & Chr$(10)
    mysql = mysql & "where Nombre" & " like '" & my_busqueda & "'" & Chr$(10)
    mysql = mysql & "OR Codigo" & " like '" & my_busqueda & "'" & Chr$(10)
    mysql = mysql & "ORDER BY Nombre ASC" & Chr$(10)

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
                ReDim Preserve my_carga_struc_proveedor(UBound(my_carga_struc_proveedor) + 1)

            End If

            If mytablex.Fields("Nombre") <> "" Then
                my_carga_struc_proveedor(UBound(my_carga_struc_proveedor)).nombre = mytablex.Fields("nombre")
            Else
                my_carga_struc_proveedor(UBound(my_carga_struc_proveedor)).nombre = ""

            End If
       
            If mytablex.Fields("codigo") <> "" Then
                my_carga_struc_proveedor(UBound(my_carga_struc_proveedor)).codigo = mytablex.Fields("codigo")
            Else
                my_carga_struc_proveedor(UBound(my_carga_struc_proveedor)).codigo = ""

            End If
      
            k = k + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

    Exit Function
carga_tipoDocumento:
    MsgBox "carga_tipoDocumento " & error$, 48, "Aviso"
    Exit Function

End Function

Public Function forma_pago(my_local As String, _
                           my_tipo As String, _
                           my_serie As String, _
                           my_numero As String, _
                           k As Integer, _
                           salida As Boolean, _
                           my_struc_fpago() As struc_fpago, _
                           acu)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo forma_pago_err

    ReDim my_struc_fpago(0)

    mysql = " select fv.LOCAL,fv.TIPO,fv.SERIE, " & Chr$(10)
    mysql = mysql & "fv.NUMERO,fv.FPAGO,fp.DESCRIPCIO," & Chr$(10)
    mysql = mysql & "fv.MONEDA,fv.TOTAL," & Chr$(10)

    If acu = "V" Or acu = "I" Or acu = "T" Then
        mysql = mysql & "FV.RECIBE,FV.SALDOS" & Chr$(10)
    ElseIf acu = "R" Or acu = "C" Then
        mysql = mysql & "FV.acuenta , FV.adetotal" & Chr$(10)

    End If

    If acu = "V" Or acu = "I" Or acu = "T" Then
        mysql = mysql & "from fpagov fv, " & Chr$(10)
    ElseIf acu = "R" Then
        mysql = mysql & "from CORDENC fv," & Chr$(10)
    ElseIf acu = "C" Then
        mysql = mysql & "from Factura fv," & Chr$(10)

    End If
  
    mysql = mysql & "fpago fp" & Chr$(10)
    mysql = mysql & "where fv.local='" & my_local & "'" & Chr$(10)
    mysql = mysql & "and fv.tipo = '" & my_tipo & "'" & Chr$(10)
    mysql = mysql & "and fv.serie = '" & my_serie & "'" & Chr$(10)
    mysql = mysql & "and fv.numero = '" & my_numero & "'" & Chr$(10)
    mysql = mysql & "and fv.fpago = fp.fpago" & Chr$(10)
   
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
                ReDim Preserve my_struc_fpago(UBound(my_struc_fpago) + 1)

            End If

            If mytablex.Fields("local") <> "" Then
                my_struc_fpago(UBound(my_struc_fpago)).local = mytablex.Fields("local")
            Else
                my_struc_fpago(UBound(my_struc_fpago)).local = ""

            End If
       
            If mytablex.Fields("tipo") <> "" Then
                my_struc_fpago(UBound(my_struc_fpago)).tipo = mytablex.Fields("tipo")
            Else
                my_struc_fpago(UBound(my_struc_fpago)).tipo = ""

            End If
       
            If mytablex.Fields("serie") <> "" Then
                my_struc_fpago(UBound(my_struc_fpago)).serie = mytablex.Fields("serie")
            Else
                my_struc_fpago(UBound(my_struc_fpago)).serie = ""

            End If
       
            If mytablex.Fields("numero") <> "" Then
                my_struc_fpago(UBound(my_struc_fpago)).Numero = mytablex.Fields("numero")
            Else
                my_struc_fpago(UBound(my_struc_fpago)).Numero = ""

            End If
       
            If mytablex.Fields("fpago") <> "" Then
                my_struc_fpago(UBound(my_struc_fpago)).fpago = mytablex.Fields("fpago")
            Else
                my_struc_fpago(UBound(my_struc_fpago)).fpago = ""

            End If
       
            If mytablex.Fields("DESCRIPCIO") <> "" Then
                my_struc_fpago(UBound(my_struc_fpago)).descripcion = mytablex.Fields("DESCRIPCIO")
            Else
                my_struc_fpago(UBound(my_struc_fpago)).descripcion = ""

            End If
       
            If mytablex.Fields("moneda") <> "" Then
                If mytablex.Fields("moneda") = "S" Then
                    my_struc_fpago(UBound(my_struc_fpago)).moneda = "Soles"
                ElseIf mytablex.Fields("moneda") = "D" Then
                    my_struc_fpago(UBound(my_struc_fpago)).moneda = "Dolares"

                End If

            Else
                my_struc_fpago(UBound(my_struc_fpago)).moneda = ""

            End If
       
            If mytablex.Fields("total") <> "" Then
                my_struc_fpago(UBound(my_struc_fpago)).total = mytablex.Fields("total")
            Else
                my_struc_fpago(UBound(my_struc_fpago)).total = 0

            End If
      
            If acu = "V" Then
                If mytablex.Fields("RECIBE") <> "" Then
                    my_struc_fpago(UBound(my_struc_fpago)).entrega = mytablex.Fields("RECIBE")
                Else
                    my_struc_fpago(UBound(my_struc_fpago)).entrega = 0

                End If
       
                If mytablex.Fields("saldos") <> "" Then
                    my_struc_fpago(UBound(my_struc_fpago)).saldos = mytablex.Fields("saldos")
                Else
                    my_struc_fpago(UBound(my_struc_fpago)).saldos = 0

                End If

            End If
      
            If acu = "R" Or acu = "C" Then
                If mytablex.Fields("acuenta") <> "" Then
                    my_struc_fpago(UBound(my_struc_fpago)).entrega = mytablex.Fields("acuenta")
                Else
                    my_struc_fpago(UBound(my_struc_fpago)).entrega = 0

                End If
       
                If mytablex.Fields("adetotal") <> "" Then
                    my_struc_fpago(UBound(my_struc_fpago)).saldos = mytablex.Fields("adetotal")
                Else
                    my_struc_fpago(UBound(my_struc_fpago)).saldos = 0

                End If

            End If
      
            k = k + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

    Exit Function
forma_pago_err:
    MsgBox "aviso en forma_pago " & error$, 48, "Aviso"
    Exit Function

End Function

Public Function ini_grid_proveedor(my_grid As MSFlexGrid)

    'utilizo per eliminare_fattura<--voce
    my_grid.Clear
    'inizializzazione Grid
    my_grid.rows = 2
    my_grid.FixedRows = 1
    my_grid.Cols = 2 'aqui se aumenta
    my_grid.FixedCols = 0

    my_grid.Row = 0
    my_grid.Col = N_colm_descripcion 'descripcion
    my_grid.Text = "Descripcion"
    my_grid.ColWidth(my_grid.Col) = 3600
    my_grid.FixedAlignment(my_grid.Col) = 0
    my_grid.ColAlignment(my_grid.Col) = 0

    my_grid.Row = 0
    my_grid.Col = N_colm_codProducto 'codProducto
    my_grid.Text = "Cod.Producto"
    my_grid.ColWidth(my_grid.Col) = 2200
    my_grid.FixedAlignment(my_grid.Col) = 0
    my_grid.ColAlignment(my_grid.Col) = 0

    Exit Function

End Function

Public Function cargar_proveedor(my_grid As MSFlexGrid, _
                                 my_carga_struc_proveedor() As struc_proveedor, _
                                 k As Integer)

    For I = 0 To k - 1
        my_grid.AddItem ""
        my_grid.Row = my_grid.rows - 1

        my_grid.Col = N_colm_descripcion
        my_grid.Text = my_carga_struc_proveedor(I).nombre

        my_grid.Col = N_colm_codProducto
        my_grid.Text = my_carga_struc_proveedor(I).codigo

    Next I

    my_grid.Row = 1
    my_grid.Col = 0

    If my_grid.Text = "" Then
        my_grid.RemoveItem 1

    End If

    Exit Function

End Function

Public Sub graba_estado_cotizacion(my_tipo As String, _
                                   my_serie As String, _
                                   my_numero As String)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    mysql = "update " & cgusuario & Chr$(10)
    mysql = mysql & "set estado ='2'" & Chr$(10)
    mysql = mysql & "where tipo='" & my_tipo & "'" & Chr$(10)
    mysql = mysql & "and serie='" & my_serie & "'" & Chr$(10)
    mysql = mysql & "and numero='" & my_numero & "'" & Chr$(10)
    cn.Execute (mysql)

End Sub

'inicio 17/08/2017 pll
Public Function ini_grid_bus_transporte(my_grid As MSFlexGrid)

    'utilizo per eliminare_fattura<--voce
    my_grid.Clear
    'inizializzazione Grid
    my_grid.rows = 2
    my_grid.FixedRows = 1
    my_grid.Cols = 2 'aqui se aumenta
    my_grid.FixedCols = 0

    my_grid.Row = 0
    my_grid.Col = N_c_nombre 'dflag
    my_grid.Text = "Nombre"
    my_grid.ColWidth(my_grid.Col) = 2200
    my_grid.FixedAlignment(my_grid.Col) = 0
    my_grid.ColAlignment(my_grid.Col) = 0

    my_grid.Row = 0
    my_grid.Col = N_c_codigo 'yausado
    my_grid.Text = "Codigo"
    my_grid.ColWidth(my_grid.Col) = 1200
    my_grid.FixedAlignment(my_grid.Col) = 0
    my_grid.ColAlignment(my_grid.Col) = 0

    Exit Function

End Function

Public Function carga_busca_transporte(my_grid As MSFlexGrid, _
                                       my_struc_transporte() As struc_transporte, _
                                       k As Integer)

    For I = 0 To k - 1
        'c = c + 1

        my_grid.AddItem ""
        my_grid.Row = my_grid.rows - 1

        my_grid.Col = N_c_nombre
        my_grid.Text = my_struc_transporte(I).nombre

        my_grid.Col = N_c_codigo
        my_grid.Text = my_struc_transporte(I).codigo

    Next I

    my_grid.Row = 1
    my_grid.Col = 0

    If my_grid.Text = "" Then
        my_grid.RemoveItem 1

    End If

    Exit Function

End Function

Public Function b_transporte(my_transporte As String, _
                             my_struc_transporte() As struc_transporte, _
                             k As Integer, _
                             salida As Boolean)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo forma_pago_err

    ReDim my_struc_transporte(0)

    mysql = "select nombre,codigo " & Chr$(10)
    mysql = mysql & "from transpor " & Chr$(10)
    mysql = mysql & "where Nombre" & " like '" & my_transporte & "'" & Chr$(10)
    mysql = mysql & "OR Codigo" & " like '" & my_transporte & "'" & Chr$(10)
    mysql = mysql & "order by NOMBRE asc" & Chr$(10)
   
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
                ReDim Preserve my_struc_transporte(UBound(my_struc_transporte) + 1)

            End If
      
            If mytablex.Fields("nombre") <> "" Then
                my_struc_transporte(UBound(my_struc_transporte)).nombre = mytablex.Fields("nombre")
            Else
                my_struc_transporte(UBound(my_struc_transporte)).nombre = ""

            End If
       
            If mytablex.Fields("codigo") <> "" Then
                my_struc_transporte(UBound(my_struc_transporte)).codigo = mytablex.Fields("codigo")
            Else
                my_struc_transporte(UBound(my_struc_transporte)).codigo = ""

            End If
          
            k = k + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

    Exit Function
forma_pago_err:
    MsgBox "aviso en forma_pago " & error$, 48, "Aviso"
    Exit Function

End Function

'**inicio 18/08/2017 pll
Public Sub b_fpago(my_struc_fpago() As struc_fpago, salida As Boolean, k As Integer)
           
    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    ReDim my_struc_fpago(0)

    mysql = "SELECT descripcio" & Chr$(10)
    mysql = mysql & "FROM fpago" & Chr$(10)
    mysql = mysql & "order by FPAGO asc" & Chr$(10)

    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF Then
        salida = False
        Exit Sub
    Else
        salida = True
        mytablex.MoveFirst
        k = 0

        Do Until mytablex.EOF

            'para cargar la msgrid
            If k > 0 Then
                ReDim Preserve my_struc_fpago(UBound(my_struc_fpago) + 1)

            End If

            If mytablex.Fields("descripcio") <> "" Then
                my_struc_fpago(UBound(my_struc_fpago)).descripcion = mytablex.Fields("descripcio")
            Else
                my_struc_fpago(UBound(my_struc_fpago)).descripcion = ""

            End If
            
            k = k + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

    Exit Sub

End Sub

Public Sub b_codigo_fpago(my_busca As String, my_codfpago As String)
           
    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    mysql = "SELECT fpago" & Chr$(10)
    mysql = mysql & "FROM fpago" & Chr$(10)
    mysql = mysql & "where descripcio ='" & my_busca & "'" & Chr$(10)

    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF Then
        salida = False
        Exit Sub
    Else
        salida = True
        mytablex.MoveFirst
        k = 0

        Do Until mytablex.EOF
            my_codfpago = mytablex.Fields("fpago")
            k = k + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

    Exit Sub

End Sub

Public Sub b_ffpago(my_serie As String, _
                    my_numero As String, _
                    my_codfpago As String, _
                    my_moneda As String, _
                    my_local As String, _
                    acu)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    mysql = "SELECT fpago,moneda,local" & Chr$(10)

    If acu = "R" Then
        mysql = mysql & "FROM CORDENC" & Chr$(10)

    End If

    If acu = "C" Or acu = "V" Or acu = "T" Or acu = "S" Or acu = "E" Or acu = "F" Then
        mysql = mysql & "FROM Factura" & Chr$(10)

    End If

    If acu = "H" Then 'cotizacion
        mysql = mysql & "FROM CCOTIZAV" & Chr$(10)

    End If

    If acu = "I" Then 'pedido
        mysql = mysql & "FROM cpedidov" & Chr$(10)

    End If

    If acu = "Z" Then 'traslado
        mysql = mysql & "FROM CTRASLAD" & Chr$(10)

    End If

    If acu = "Q" Then 'traslado
        mysql = mysql & "FROM CREQUISA" & Chr$(10)

    End If

    mysql = mysql & "where serie ='" & my_serie & "'" & Chr$(10)
    mysql = mysql & "and numero ='" & my_numero & "'" & Chr$(10)

    If acu = "C" Or acu = "S" Then 'compras
        mysql = mysql & "and tipo ='" & my_tipo & "'" & Chr$(10)

    End If

    If acu = "T" Then
        mysql = mysql & "and tipo ='" & my_tipo & "'" & Chr$(10)

    End If

    If acu = "Z" Then
        mysql = mysql & "and tipo ='" & acu & "'" & Chr$(10)

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

            If mytablex.Fields("fpago") = Null Then
                my_codfpago = ""
            Else

                If mytablex.Fields("fpago") <> Null Then
                    my_codfpago = mytablex.Fields("fpago")
                Else
                    my_codfpago = mytablex.Fields("fpago")

                End If

                my_moneda = mytablex.Fields("moneda")
                my_local = mytablex.Fields("local")

            End If

            k = k + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

    Exit Sub

End Sub

'**
Public Sub b_descri_fpago(my_busca As String, my_descri As String)
           
    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    mysql = "SELECT descripcio" & Chr$(10)
    mysql = mysql & "FROM fpago" & Chr$(10)
    mysql = mysql & "where fpago ='" & my_busca & "'" & Chr$(10)

    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF Then
        salida = False
        Exit Sub
    Else
        salida = True
        mytablex.MoveFirst
        k = 0

        Do Until mytablex.EOF
            my_descri = mytablex.Fields("descripcio")
            k = k + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

    Exit Sub

End Sub

Public Function b_codcliente(my_busqueda As String, _
                             my_codcliente As String, _
                             salida As Boolean)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo b_codcliente

    mysql = "select Codigo from clientes" & Chr$(10)
    mysql = mysql & "where Nombre" & " like '" & my_busqueda & "'" & Chr$(10)

    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF Then
        salida = False
        Exit Function
    Else
        salida = True
        my_codcliente = mytablex.Fields("Codigo")

    End If

    mytablex.Close

    Exit Function

b_codcliente:
    MsgBox "b_codcliente" & error$, 48, "Aviso"
    Exit Function

End Function

Public Function c_vendedor(my_struc_vendedor() As struc_vendedor, _
                           salida As Boolean, _
                           k As Integer)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo carga_vendedor

    ReDim my_struc_vendedor(0)

    mysql = "select nombre,codigo  from vendedor" & Chr$(10)

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
                ReDim Preserve my_struc_vendedor(UBound(my_struc_vendedor) + 1)

            End If

            If mytablex.Fields("nombre") <> "" Then
                my_struc_vendedor(UBound(my_struc_vendedor)).nombre = mytablex.Fields("nombre")
            Else
                my_struc_vendedor(UBound(my_struc_vendedor)).nombre = ""

            End If
       
            If mytablex.Fields("codigo") <> "" Then
                my_struc_vendedor(UBound(my_struc_vendedor)).codigo = mytablex.Fields("codigo")
            Else
                my_struc_vendedor(UBound(my_struc_vendedor)).codigo = ""

            End If
      
            k = k + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

    Exit Function
carga_vendedor:
    MsgBox "carga_vendedor " & error$, 48, "Aviso"
    Exit Function

End Function

Public Function configura_moneda(my_vdolar As String, salida As Boolean)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo configura_moneda

    mysql = "select vdolar" & Chr$(10)
    mysql = mysql & "from parame" & Chr$(10)

    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF Then
        salida = False
        Exit Function
    Else
        salida = True
        my_vdolar = mytablex.Fields("vdolar")

    End If

    mytablex.Close

    Exit Function
configura_moneda:
    MsgBox "configura_moneda " & error$, 48, "Aviso"
    Exit Function

End Function

Public Sub b_producto_precio(my_codProducto As String, my_costou As String)
           
    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    ReDim my_carga_busca_producto(0)

    mysql = "SELECT costou,igv" & Chr$(10)
    mysql = mysql & "From producto" & Chr$(10)
    mysql = mysql & "where producto='" & my_codProducto & "'" & Chr$(10)

    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF Then
        salida = False
        Exit Sub
    Else
        salida = True
        my_costou = mytablex.Fields("costou")

    End If

    mytablex.Close

    Exit Sub

End Sub

