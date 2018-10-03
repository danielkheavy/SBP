VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form TUNIFLEX 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Interface Uniflex"
   ClientHeight    =   8520
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   14985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Productos"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin MSDataGridLib.DataGrid dbGrid1 
      Height          =   7455
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   13150
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label ii 
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin VB.Menu flo34 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "TUNIFLEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cnu As New ADODB.Connection

Private Sub Command1_Click()

    On Error GoTo cmd8912_err

    cnu.CursorLocation = adUseClient
    cnu.CommandTimeout = 1024
    cnu.Open "Driver={SQL Server};Server=" & menup.vservidor & ";Database=axes"
    graba_producto
    'graba_marca
    'graba_familia
    'graba_precios
    'visualiza
    'ver_sql
 
    Exit Sub
cmd8912_err:
    MsgBox "Conexion ha fallado " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub ver_sql()

    Dim buf      As String

    Dim mytabley As New ADODB.Recordset

    buf = "SELECT     TOP (100) PERCENT dbo.Producto.idmarca ,dbo.ProductoServicio.PKID AS IDProducto, dbo.ProductoServicio.Codigo, dbo.ProductoServicio.Descripcion AS Producto, dbo.Unidad.Abreviacion,"
    buf = buf & "                      dbo.Unidad.Factor, dbo.Escala.Desde, dbo.Escala.Hasta, dbo.Escala.PrecioIncImpto, dbo.Unidad.Referencia, dbo.Producto.CostoPromedioSoles,"
    buf = buf & "                       CASE dbo.ImpuestoProductoServicio.Inafecto WHEN 0 THEN ((dbo.Producto.CostoPromedioSoles * dbo.Unidad.Factor) * 1.18)"
    buf = buf & "                       ELSE (dbo.Producto.CostoPromedioSoles * dbo.Unidad.Factor) END AS CPR,"
    buf = buf & "                       CASE dbo.ImpuestoProductoServicio.Inafecto WHEN 0 THEN ((dbo.Producto.CostoUltimaCompraSoles * 1.18) * dbo.Unidad.Factor)"
    buf = buf & "                       ELSE (dbo.Producto.CostoUltimaCompraSoles * dbo.Unidad.Factor) END AS CUC, dbo.Sucursal.PKID AS IDSucursal, dbo.ItemListaPrecios.Desactivado AS Activo,"
    buf = buf & "                       dbo.ClaseProductoServicio.PKID AS IDLinea,  dbo.ListaPrecios.Descripcion, dbo.CategoriaCliente.PKID AS IDCategoria,"
    buf = buf & "                       dbo.ListaPrecios.PKID AS IDListaPrecio"
    buf = buf & " FROM         dbo.ListaPrecios INNER JOIN"
    buf = buf & "                       dbo.ItemListaPrecios ON dbo.ListaPrecios.PKID = dbo.ItemListaPrecios.IDListaPrecios INNER JOIN"
    buf = buf & "                       dbo.SucursalListaPrecios ON dbo.ListaPrecios.PKID = dbo.SucursalListaPrecios.IDListaPrecios INNER JOIN"
    buf = buf & "                       dbo.Sucursal ON dbo.SucursalListaPrecios.IDSucursal = dbo.Sucursal.PKID INNER JOIN"
    buf = buf & "                       dbo.UnidadItemListaPrecios ON dbo.ItemListaPrecios.PKID = dbo.UnidadItemListaPrecios.IDItemListaPrecios INNER JOIN"
    buf = buf & "                       dbo.ProductoServicio ON dbo.ItemListaPrecios.IDProducto = dbo.ProductoServicio.PKID INNER JOIN"
    buf = buf & "                       dbo.Escala ON dbo.UnidadItemListaPrecios.PKID = dbo.Escala.IDUnidadItemListaPrecios INNER JOIN"
    buf = buf & "                       dbo.Unidad ON dbo.UnidadItemListaPrecios.IDUnidad = dbo.Unidad.PKID INNER JOIN"
    buf = buf & "                       dbo.Producto ON dbo.ProductoServicio.PKID = dbo.Producto.PKID AND dbo.ItemListaPrecios.IDProducto = dbo.Producto.PKID INNER JOIN"
    buf = buf & "                      dbo.ImpuestoProductoServicio ON dbo.ProductoServicio.PKID = dbo.ImpuestoProductoServicio.IDProductoServicio INNER JOIN"
    buf = buf & "                      dbo.ClaseProductoServicio ON dbo.ProductoServicio.IDClaseProductoServicio = dbo.ClaseProductoServicio.PKID INNER JOIN"
    buf = buf & "                       dbo.Persona ON dbo.ListaPrecios.IDCliente = dbo.Persona.PKID INNER JOIN"
    buf = buf & "                       dbo.CategoriaCliente ON dbo.ListaPrecios.IDCategoriaCliente = dbo.CategoriaCliente.PKID"
    buf = buf & " Where (dbo.Escala.EsVigente = 1 AND dbo.ProductoServicio.PKID=10145)"

    mytabley.Open buf, cnu, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = mytabley
 
End Sub

Sub visualiza()

    Dim mytablex  As New ADODB.Recordset

    Dim mytabley  As New ADODB.Recordset

    Dim mytablez  As New ADODB.Recordset

    Dim mytablezz As New ADODB.Recordset

    mytabley.Open "select * from itemlistaprecios ", cnu, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = mytabley
 
    Label1 = "itemlistaprecios"

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from escala ", cnu, adOpenStatic, adLockOptimistic
    Set DBGrid2.DataSource = mytablex
    Label2 = "escala"
    mytablez.Open "select * from UNIDADITEMLISTAPRECIOS ", cnu, adOpenStatic, adLockOptimistic
    Set dbgrid3.DataSource = mytablez
    Label3 = "unidaditemlistaprecios"

    mytablezz.Open "select * from UNIDAD ", cnu, adOpenStatic, adLockOptimistic
    Set DBGrid4.DataSource = mytablezz
    Label4 = "unidad"

End Sub

Sub graba_precios()

    Dim buf       As String

    Dim mytablex  As New ADODB.Recordset

    Dim mytabley  As New ADODB.Recordset

    Dim mytablez  As New ADODB.Recordset

    Dim mytablezz As New ADODB.Recordset

    ReDim xprecio(10) As Double
    ReDim xunidad(10) As String
    ReDim xfactor(10) As Double

    Dim xlocal    As String

    Dim xproducto As String

    Dim I         As Integer

    Dim vr

    On Error GoTo cm8912_err

    buf = "SELECT     TOP (100) PERCENT dbo.ProductoServicio.PKID AS IDProducto, dbo.ProductoServicio.Codigo, dbo.ProductoServicio.Descripcion AS Producto, dbo.Unidad.Abreviacion,"
    buf = buf & "                      dbo.Unidad.Factor, dbo.Escala.Desde, dbo.Escala.Hasta, dbo.Escala.PrecioIncImpto, dbo.Unidad.Referencia, dbo.Producto.CostoPromedioSoles,"
    buf = buf & "                       CASE dbo.ImpuestoProductoServicio.Inafecto WHEN 0 THEN ((dbo.Producto.CostoPromedioSoles * dbo.Unidad.Factor) * 1.18)"
    buf = buf & "                       ELSE (dbo.Producto.CostoPromedioSoles * dbo.Unidad.Factor) END AS CPR,"
    buf = buf & "                       CASE dbo.ImpuestoProductoServicio.Inafecto WHEN 0 THEN ((dbo.Producto.CostoUltimaCompraSoles * 1.18) * dbo.Unidad.Factor)"
    buf = buf & "                       ELSE (dbo.Producto.CostoUltimaCompraSoles * dbo.Unidad.Factor) END AS CUC, dbo.Sucursal.PKID AS IDSucursal, dbo.ItemListaPrecios.Desactivado AS Activo,"
    buf = buf & "                       dbo.ClaseProductoServicio.PKID AS IDLinea,  dbo.ListaPrecios.Descripcion, dbo.CategoriaCliente.PKID AS IDCategoria,"
    buf = buf & "                       dbo.ListaPrecios.PKID AS IDListaPrecio"
    buf = buf & " FROM         dbo.ListaPrecios INNER JOIN"
    buf = buf & "                       dbo.ItemListaPrecios ON dbo.ListaPrecios.PKID = dbo.ItemListaPrecios.IDListaPrecios INNER JOIN"
    buf = buf & "                       dbo.SucursalListaPrecios ON dbo.ListaPrecios.PKID = dbo.SucursalListaPrecios.IDListaPrecios INNER JOIN"
    buf = buf & "                       dbo.Sucursal ON dbo.SucursalListaPrecios.IDSucursal = dbo.Sucursal.PKID INNER JOIN"
    buf = buf & "                       dbo.UnidadItemListaPrecios ON dbo.ItemListaPrecios.PKID = dbo.UnidadItemListaPrecios.IDItemListaPrecios INNER JOIN"
    buf = buf & "                       dbo.ProductoServicio ON dbo.ItemListaPrecios.IDProducto = dbo.ProductoServicio.PKID INNER JOIN"
    buf = buf & "                       dbo.Escala ON dbo.UnidadItemListaPrecios.PKID = dbo.Escala.IDUnidadItemListaPrecios INNER JOIN"
    buf = buf & "                       dbo.Unidad ON dbo.UnidadItemListaPrecios.IDUnidad = dbo.Unidad.PKID INNER JOIN"
    buf = buf & "                       dbo.Producto ON dbo.ProductoServicio.PKID = dbo.Producto.PKID AND dbo.ItemListaPrecios.IDProducto = dbo.Producto.PKID INNER JOIN"
    buf = buf & "                      dbo.ImpuestoProductoServicio ON dbo.ProductoServicio.PKID = dbo.ImpuestoProductoServicio.IDProductoServicio INNER JOIN"
    buf = buf & "                      dbo.ClaseProductoServicio ON dbo.ProductoServicio.IDClaseProductoServicio = dbo.ClaseProductoServicio.PKID INNER JOIN"
    buf = buf & "                       dbo.Persona ON dbo.ListaPrecios.IDCliente = dbo.Persona.PKID INNER JOIN"
    buf = buf & "                       dbo.CategoriaCliente ON dbo.ListaPrecios.IDCategoriaCliente = dbo.CategoriaCliente.PKID"
    buf = buf & " Where (dbo.Escala.EsVigente = 1)"

    cn.Execute ("delete from precios ")
    MsgBox "Enter"
    mytablex.Open "select * from precios", cn, adOpenStatic, adLockOptimistic
    mytabley.Open buf, cnu, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = mytabley
    Do

        If mytabley.EOF Then Exit Do
        vr = DoEvents()
        ii = "" & I
        I = 0
        mytablez.Open "select * from  escala  where pkid='" & mytabley.Fields("pkid") & "'", cnu, adOpenStatic, adLockOptimistic

        If mytablez.RecordCount > 0 Then
            xlocal = Format(Val("" & mytabley.Fields("idlistaprecios")), "00")
            xproducto = "" & mytabley.Fields("idproducto")

            For I = 0 To 9
                xprecio(I) = 0
                xunidad(I) = ""
                xfactor(I) = 0
            Next I

            I = 0
            Do

                If mytablez.EOF Then Exit Do
                mytablezz.Open "select  unidad.abreviacion as unidad,unidad.factor as factor  from  unidaditemlistaprecios,unidad   where unidaditemlistaprecios.idunidad=unidad.pkid and unidaditemlistaprecios.idItemlistaprecios='" & mytablez.Fields("idunidaditemlistaprecios") & "'", cnu, adOpenStatic, adLockOptimistic

                If mytablezz.RecordCount > 0 Then
                    'MsgBox ""
                    I = I + 1
                    xprecio(I) = Val("" & mytablez.Fields("precio"))
                    xunidad(I) = "" & mytablezz.Fields("unidad")
                    xfactor(I) = Val("" & mytablezz.Fields("factor"))
       
                End If

                mytablezz.Close
    
                mytablez.MoveNext
            Loop

        End If

        mytablez.Close

        'MsgBox ""
        'Exit Do
        If I > 0 Then
            mytablex.AddNew
            mytablex.Fields("local") = xlocal
            mytablex.Fields("producto") = xproducto

            If xprecio(1) > 0 Then
                mytablex.Fields("pventa1") = xprecio(1)
                mytablex.Fields("unidad1") = xunidad(1)
                mytablex.Fields("factor1") = xfactor(1)

                'MsgBox ""
                'End
            End If

            If xprecio(2) > 0 Then
                mytablex.Fields("pventa2") = xprecio(2)
                mytablex.Fields("unidad2") = xunidad(2)
                mytablex.Fields("factor2") = xfactor(2)

            End If

            If xprecio(3) > 0 Then
                mytablex.Fields("pventa3") = xprecio(3)
                mytablex.Fields("unidad3") = xunidad(3)
                mytablex.Fields("factor3") = xfactor(3)

            End If

            mytablex.Update

        End If
 
        mytabley.MoveNext
    Loop
    Exit Sub
cm8912_err:
    MsgBox error$, 48, "Aviso"
    Exit Sub

End Sub

Sub graba_familia()

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    cn.Execute ("delete from familia ")
    mytablex.Open "select * from familia ", cn, adOpenStatic, adLockOptimistic
    mytabley.Open "select * from claseproductoservicio ", cnu, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = mytabley
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew
        mytablex.Fields("familia") = "" & mytabley.Fields("pkid")
        mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcion")
        mytablex.Update
        mytabley.MoveNext
    Loop
    mytablex.Close
    mytabley.Close

End Sub

Sub graba_marca()

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    cn.Execute ("delete from marca ")
    mytablex.Open "select * from marca ", cn, adOpenStatic, adLockOptimistic
    mytabley.Open "select * from marca ", cnu, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = mytabley
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew
        mytablex.Fields("marca") = "" & mytabley.Fields("pkid")
        mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcion")
        mytablex.Update
        mytabley.MoveNext
    Loop
    mytablex.Close
    mytabley.Close

End Sub

Sub graba_producto()

    Dim buf As String

    Dim I   As Long

    Dim vr

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    buf = "SELECT     TOP (100) PERCENT dbo.ProductoServicio.PKID AS IDProducto,dbo.producto.idmarca, dbo.ProductoServicio.Codigo, dbo.ProductoServicio.Descripcion AS Producto, dbo.Unidad.Abreviacion as unidad,"
    buf = buf & "                       dbo.Unidad.Factor, dbo.Escala.Desde, dbo.Escala.Hasta, dbo.Escala.PrecioIncImpto, dbo.Unidad.Referencia, dbo.Producto.CostoPromedioSoles,"
    buf = buf & "                       CASE dbo.ImpuestoProductoServicio.Inafecto WHEN 0 THEN ((dbo.Producto.CostoPromedioSoles * dbo.Unidad.Factor) * 1.18)"
    buf = buf & "                       ELSE (dbo.Producto.CostoPromedioSoles * dbo.Unidad.Factor) END AS CPR,"
    buf = buf & "                       CASE dbo.ImpuestoProductoServicio.Inafecto WHEN 0 THEN ((dbo.Producto.CostoUltimaCompraSoles * 1.18) * dbo.Unidad.Factor)"
    buf = buf & "                       ELSE (dbo.Producto.CostoUltimaCompraSoles * dbo.Unidad.Factor) END AS CUC, dbo.Sucursal.PKID AS IDSucursal, dbo.ItemListaPrecios.Desactivado AS Activo,"
    buf = buf & "                       dbo.ClaseProductoServicio.PKID AS IDLinea,  dbo.ListaPrecios.Descripcion, dbo.CategoriaCliente.PKID AS IDCategoria,"
    buf = buf & "                       dbo.ListaPrecios.PKID AS IDListaPrecio"
    buf = buf & " FROM         dbo.ListaPrecios INNER JOIN"
    buf = buf & "                       dbo.ItemListaPrecios ON dbo.ListaPrecios.PKID = dbo.ItemListaPrecios.IDListaPrecios INNER JOIN"
    buf = buf & "                       dbo.SucursalListaPrecios ON dbo.ListaPrecios.PKID = dbo.SucursalListaPrecios.IDListaPrecios INNER JOIN"
    buf = buf & "                       dbo.Sucursal ON dbo.SucursalListaPrecios.IDSucursal = dbo.Sucursal.PKID INNER JOIN"
    buf = buf & "                       dbo.UnidadItemListaPrecios ON dbo.ItemListaPrecios.PKID = dbo.UnidadItemListaPrecios.IDItemListaPrecios INNER JOIN"
    buf = buf & "                       dbo.ProductoServicio ON dbo.ItemListaPrecios.IDProducto = dbo.ProductoServicio.PKID INNER JOIN"
    buf = buf & "                       dbo.Escala ON dbo.UnidadItemListaPrecios.PKID = dbo.Escala.IDUnidadItemListaPrecios INNER JOIN"
    buf = buf & "                       dbo.Unidad ON dbo.UnidadItemListaPrecios.IDUnidad = dbo.Unidad.PKID INNER JOIN"
    buf = buf & "                       dbo.Producto ON dbo.ProductoServicio.PKID = dbo.Producto.PKID AND dbo.ItemListaPrecios.IDProducto = dbo.Producto.PKID INNER JOIN"
    buf = buf & "                       dbo.ImpuestoProductoServicio ON dbo.ProductoServicio.PKID = dbo.ImpuestoProductoServicio.IDProductoServicio INNER JOIN"
    buf = buf & "                       dbo.ClaseProductoServicio ON dbo.ProductoServicio.IDClaseProductoServicio = dbo.ClaseProductoServicio.PKID INNER JOIN"
    buf = buf & "                       dbo.Persona ON dbo.ListaPrecios.IDCliente = dbo.Persona.PKID INNER JOIN"
    buf = buf & "                       dbo.CategoriaCliente ON dbo.ListaPrecios.IDCategoriaCliente = dbo.CategoriaCliente.PKID"
    buf = buf & " Where (dbo.Escala.EsVigente = 1 )"
    'buf = buf & " Where (dbo.Escala.EsVigente = 1 AND dbo.ProductoServicio.PKID=10145)"

    I = 0
    cn.Execute ("delete from precios ")
    cn.Execute ("delete from producto ")

    mytabley.Open buf, cnu, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = mytabley
    Do

        If mytabley.EOF Then Exit Do
        vr = DoEvents()
        I = I + 1
        Label1 = "" & I

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open "select * from producto where producto='" & mytabley.Fields("idproducto") & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.AddNew
            mytablex.Fields("producto") = "" & mytabley.Fields("idproducto")
            mytablex.Fields("barras") = ""
            mytablex.Fields("descripcio") = Mid$("" & mytabley.Fields("producto"), 1, 80)
            mytablex.Fields("descorto") = Mid$("" & mytabley.Fields("producto"), 1, 20)
            mytablex.Fields("presenta") = ""
            mytablex.Fields("familia") = "" & mytabley.Fields("idlinea")
            mytablex.Fields("subfamilia") = ""
            mytablex.Fields("seccion") = ""
            mytablex.Fields("marca") = "" & mytabley.Fields("idmarca")
            mytablex.Fields("categoria") = "" & mytabley.Fields("idcategoria")
            mytablex.Fields("linea") = ""
            mytablex.Fields("color") = ""
            mytablex.Fields("fabrica") = ""
            mytablex.Fields("serie") = ""
            mytablex.Fields("peso") = ""
            mytablex.Fields("servicio") = ""
            mytablex.Fields("vecaja") = "S"
            mytablex.Fields("igv") = 18
            mytablex.Fields("isc") = 0
            mytablex.Fields("pesokgr") = 0.001
            mytablex.Fields("comision") = 0
            mytablex.Fields("monedac") = "S"
            mytablex.Fields("unidad") = "UND"
            mytablex.Fields("factor") = 1
            mytablex.Fields("costou") = 0
            mytablex.Fields("costop") = 0
            mytablex.Fields("monedav") = "S"
            'mytablex.Fields("unidad") = "" & mytabley.Fields("unidad")
            mytablex.Update

        End If

        '------------ahora grabamos los precios
        mytablez.Open "select * from precios where local='" & Format(Trim("" & mytabley.Fields("idlistaprecio")), "00") & "' and producto='" & Trim("" & mytabley.Fields("idproducto")) & "' ", cn, adOpenStatic, adLockOptimistic

        If mytablez.RecordCount = 0 Then
            mytablez.AddNew

        End If

        mytablez.Fields("local") = Format(Trim("" & mytabley.Fields("idlistaprecio")), "00")
        mytablez.Fields("producto") = "" & mytabley.Fields("idproducto")

        If Val("" & mytablez.Fields("pventa1")) = 0 Then
            mytablez.Fields("pventa1") = Val("" & mytabley.Fields("PrecioIncImpto"))
            mytablez.Fields("unidad1") = "" & mytabley.Fields("unidad")
            mytablez.Fields("factor1") = Val("" & mytabley.Fields("factor"))
            GoTo amigo

        End If

        If Val("" & mytablez.Fields("pventa2")) = 0 Then
            mytablez.Fields("pventa2") = Val("" & mytabley.Fields("PrecioIncImpto"))
            mytablez.Fields("unidad2") = "" & mytabley.Fields("unidad")
            mytablez.Fields("factor2") = Val("" & mytabley.Fields("factor"))
            GoTo amigo

        End If

        If Val("" & mytablez.Fields("pventa3")) = 0 Then
            mytablez.Fields("pventa3") = Val("" & mytabley.Fields("PrecioIncImpto"))
            mytablez.Fields("unidad3") = "" & mytabley.Fields("unidad")
            mytablez.Fields("factor3") = Val("" & mytabley.Fields("factor"))
            GoTo amigo

        End If

        If Val("" & mytablez.Fields("pventa4")) = 0 Then
            mytablez.Fields("pventa4") = Val("" & mytabley.Fields("PrecioIncImpto"))
            mytablez.Fields("unidad4") = "" & mytabley.Fields("unidad")
            mytablez.Fields("factor4") = Val("" & mytabley.Fields("factor"))
            GoTo amigo

        End If

        If Val("" & mytablez.Fields("pventa5")) = 0 Then
            mytablez.Fields("pventa5") = Val("" & mytabley.Fields("PrecioIncImpto"))
            mytablez.Fields("unidad5") = "" & mytabley.Fields("unidad")
            mytablez.Fields("factor5") = Val("" & mytabley.Fields("factor"))
            GoTo amigo

        End If

amigo:
        mytablez.Update
        mytablez.Close
        mytablex.Close
    
        mytabley.MoveNext
    Loop
    mytabley.Close

End Sub

Private Sub flo34_Click()
    TUNIFLEX.Hide
    Unload TUNIFLEX

End Sub

