Attribute VB_Name = "Mexcel"

'inicio 16/06/2017 pll
Public objWorkBook  As Object

Public objWorksheet As Object

Global oXL          As Object

'inicio 01/09/2017 pll
'Public objExcel                                     As Excel.Application
Public xlibro       As Excel.Workbook

Type struc_producto

    familia                         As String
    producto                        As String
    descripcion                     As String
    xcanti                          As Double
    xcostou                         As Double
    xtotal                          As Double
    moneda                          As String
    unidad                          As String
    factor                          As String
    precio                          As Double

End Type

Type struc_saldo_actual

    familia                        As String
    subfamilia                     As String
    categoria                      As String
    producto                       As String
    descripcion                    As String
    unidad                         As String
    factor                         As String
    cantidad                       As Double
    total                          As String
    costop                         As Double
    costou                         As Double
    moneda                         As String
    minimo                         As String

End Type

Global my_struc_producto()     As struc_producto

Global my_struc_saldo_actual() As struc_saldo_actual

'inicio 02/08/2017 pll
Type struc_cotizacion_total_excel

    cantidad                       As Double
    total                          As Double
    subtotal                       As Double
    igv                            As Double
    producto                       As String
    descripcion                    As String
    unidad                         As String
    factor                         As String
    precio                         As Double

End Type

Global my_struc_cotizacion_total_excel() As struc_cotizacion_total_excel

'fin 02/08/2017 pll
'inicio 08/08/2017 pll
Type struc_datos_empresa

    nombre                         As String
    nombrec                        As String
    direccion                      As String
    dpto                           As String
    distrito                       As String
    telefono                       As String
    codigo1                        As String
    correo                         As String
    Toperacion                     As String
    timpresion                     As String
    esunat                         As String
  
    ' Varios Locales FE 18/05/2018
    CodSede                     As String
    ' Varios Locales FE 18/05/2018
  
End Type

Global my_struc_datos_empresa() As struc_datos_empresa

'fin 08/08/2017 pll
'inicio 17/08/2017 pll
Type struc_Etransporte

    partida                        As String
    destino                        As String
    RUC                            As String
    nombreT                        As String
    nombrec                        As String
    fecha                          As String
    placa                          As String
    licencia                       As String
    marca                          As String
    vehiculo                       As String

End Type

Global my_struc_Etransporte() As struc_Etransporte

'inicio 13/12/2017 pll registro ventas - compras
Type struc_Rventas

    num_correlativo                 As Integer
    fecha                           As String
    local                           As String
    tipo                            As String
    serie                           As String
    Numero                          As String
    acu                             As String
    codigo                          As String
    nombre                          As String
    estado                          As String
    subtotal                        As Currency
    total                           As Currency
    gravado                         As String
    tisc                            As String
    impuesto                        As Currency
    tivap                           As Currency
    percepcion                      As Currency
    servicioco                      As Currency
    tdetra                          As Currency
    paridad                         As Currency
    fechasunat                      As String
    tipo1                           As String
    serie1                          As String
    numero1                         As String
    observa                         As String

End Type

'fin 13/12/2017 pll registro ventas - compras
Global my_struc_Rventas() As struc_Rventas

'Testing Facturacion Electronica 14/03/2018
'Reemplaza xlocal para caja
'Testing Facturacion Electronica 14/03/2018
Public Function Datos_Empresa(my_struc_datos_empresa() As struc_datos_empresa, _
                              xlocal As String, _
                              salida As Boolean, _
                              k As Integer)
                              
    Dim mytablex As New ADODB.Recordset

    Dim mysql    As String

    Datos_Empresa = ok

    On Error GoTo Datos_Empresa

    ReDim my_struc_datos_empresa(0)

    ' Varios Locales FE 18/05/2018

    ' Varios Locales FE 18/05/2018
    mysql = ""
    mysql = "SELECT nombre,nombrec,direccion,dpto,distrito," & Chr$(10)
    mysql = mysql & "telefono,codigo1,correo,toperacion,esunat," & Chr$(10)
    mysql = mysql & "timpresion,codsede " & Chr$(10)
    mysql = mysql & "FROM tlocal" & Chr$(10)

    'Testing Facturacion Electronica 14/03/2018
    'mysql = mysql & "WHERE codigo= " & caja & " " & Chr$(10)
    mysql = mysql & "WHERE codigo= " & xlocal & " " & Chr$(10)
    'Testing Facturacion Electronica 14/03/2018

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
                ReDim Preserve my_struc_datos_empresa(UBound(my_struc_datos_empresa) + 1)

            End If
     
            If mytablex.Fields("nombre") <> "" Then
                my_struc_datos_empresa(UBound(my_struc_datos_empresa)).nombre = mytablex.Fields("nombre")
            Else
                my_struc_datos_empresa(UBound(my_struc_datos_empresa)).nombre = ""

            End If
      
            If mytablex.Fields("nombrec") <> "" Then
                my_struc_datos_empresa(UBound(my_struc_datos_empresa)).nombrec = mytablex.Fields("nombrec")
            Else
                my_struc_datos_empresa(UBound(my_struc_datos_empresa)).nombrec = ""

            End If
      
            If mytablex.Fields("direccion") <> "" Then
                my_struc_datos_empresa(UBound(my_struc_datos_empresa)).direccion = mytablex.Fields("direccion")
            Else
                my_struc_datos_empresa(UBound(my_struc_datos_empresa)).direccion = ""

            End If
    
            If mytablex.Fields("dpto") <> "" Then
                my_struc_datos_empresa(UBound(my_struc_datos_empresa)).dpto = mytablex.Fields("dpto")
            Else
                my_struc_datos_empresa(UBound(my_struc_datos_empresa)).dpto = ""

            End If
    
            If mytablex.Fields("distrito") <> "" Then
                my_struc_datos_empresa(UBound(my_struc_datos_empresa)).distrito = mytablex.Fields("distrito")
            Else
                my_struc_datos_empresa(UBound(my_struc_datos_empresa)).distrito = ""

            End If
   
            If mytablex.Fields("telefono") <> "" Then
                my_struc_datos_empresa(UBound(my_struc_datos_empresa)).telefono = mytablex.Fields("telefono")
            Else
                my_struc_datos_empresa(UBound(my_struc_datos_empresa)).telefono = ""

            End If
   
            If mytablex.Fields("codigo1") <> "" Then
                my_struc_datos_empresa(UBound(my_struc_datos_empresa)).codigo1 = mytablex.Fields("codigo1")
            Else
                my_struc_datos_empresa(UBound(my_struc_datos_empresa)).codigo1 = ""

            End If
   
            If mytablex.Fields("correo") <> "" Then
                my_struc_datos_empresa(UBound(my_struc_datos_empresa)).correo = mytablex.Fields("correo")
            Else
                my_struc_datos_empresa(UBound(my_struc_datos_empresa)).correo = ""

            End If
   
            If mytablex.Fields("toperacion") <> "" Then
                my_struc_datos_empresa(UBound(my_struc_datos_empresa)).Toperacion = mytablex.Fields("toperacion")
            Else
                my_struc_datos_empresa(UBound(my_struc_datos_empresa)).Toperacion = ""

            End If
   
            If mytablex.Fields("timpresion") <> "" Then
                my_struc_datos_empresa(UBound(my_struc_datos_empresa)).timpresion = mytablex.Fields("timpresion")
            Else
                my_struc_datos_empresa(UBound(my_struc_datos_empresa)).timpresion = ""

            End If
   
            If mytablex.Fields("esunat") <> "" Then
                my_struc_datos_empresa(UBound(my_struc_datos_empresa)).esunat = mytablex.Fields("esunat")
            Else
                my_struc_datos_empresa(UBound(my_struc_datos_empresa)).esunat = ""

            End If
   
            ' Varios Locales FE 18/05/2018
            If mytablex.Fields("codsede") <> "" Then
                my_struc_datos_empresa(UBound(my_struc_datos_empresa)).CodSede = mytablex.Fields("codsede")
            Else
                my_struc_datos_empresa(UBound(my_struc_datos_empresa)).CodSede = ""

            End If

            ' Varios Locales FE 18/05/2018
   
            k = k + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
    Exit Function

Datos_Empresa:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

End Function

Public Sub carga_datos_empresa(my_struc_datos_empresa() As struc_datos_empresa, _
                               k As Integer)

    On Error GoTo carga_datos_empresa

    If repdocrv.titulo = "REGISTRO DE VENTAS S/" Or repdocrv.titulo = "REGISTRO DE COMPRAS S/" Or my_report = "Formapago" Or my_report = "Detalle" Then

        For I = 0 To k - 1
            'objWorksheet.Cells(2, 7) = my_struc_datos_empresa(i).nombre
            objWorksheet.Range("G2").ColumnWidth = 12
            objWorksheet.Cells(2, 7) = my_struc_datos_empresa(I).nombre
    
            objWorksheet.Range("G3").ColumnWidth = 12
            objWorksheet.Cells(3, 7) = "Direccion : " & my_struc_datos_empresa(I).direccion
    
            objWorksheet.Range("G4").ColumnWidth = 12
            objWorksheet.Cells(4, 7) = "Departamento : " & my_struc_datos_empresa(I).dpto

            objWorksheet.Range("G5").ColumnWidth = 12
            objWorksheet.Cells(5, 7) = "Distrito: " & my_struc_datos_empresa(I).distrito
   
            objWorksheet.Range("G6").ColumnWidth = 12
            objWorksheet.Cells(6, 7) = "Telefono : " & my_struc_datos_empresa(I).telefono
   
            objWorksheet.Range("G7").ColumnWidth = 12
            objWorksheet.Cells(7, 7) = "RUC : " & my_struc_datos_empresa(I).codigo1
   
        Next I
   
    ElseIf repinv.Visible = True Or my_acu = "H" Or my_acu = "T" Or my_acu = "J" Or my_acu = "D" Or my_acu = "V" Or my_acu = "C" Or explorap.Caption = "Documentos Guia Remision Compra" Then

        For I = 0 To k - 1
            'objWorksheet.Cells(2, 7) = my_struc_datos_empresa(i).nombre
            objWorksheet.Range("D2").ColumnWidth = 12
            objWorksheet.Cells(2, 4) = my_struc_datos_empresa(I).nombre
    
            objWorksheet.Range("D3").ColumnWidth = 12
            objWorksheet.Cells(3, 4) = "Direccion : " & my_struc_datos_empresa(I).direccion
    
            objWorksheet.Range("D4").ColumnWidth = 12
            objWorksheet.Cells(4, 4) = "Departamento : " & my_struc_datos_empresa(I).dpto
    
            objWorksheet.Range("D5").ColumnWidth = 12
            objWorksheet.Cells(5, 4) = "Distrito: " & my_struc_datos_empresa(I).distrito
   
            objWorksheet.Range("D6").ColumnWidth = 12
            objWorksheet.Cells(6, 4) = "Telefono : " & my_struc_datos_empresa(I).telefono
   
            objWorksheet.Range("D7").ColumnWidth = 12
            objWorksheet.Cells(7, 4) = "RUC : " & my_struc_datos_empresa(I).codigo1
   
        Next I

    End If

    Exit Sub
 
carga_datos_empresa:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

    End Select

End Sub

Function AbreExcel() As Integer

    On Error GoTo AbreExcelERR

    On Error Resume Next

    ' Si el Excel eta abierto coje el control:
    Set oXL = GetObject(my_path, "Excel.application")

    'Set oXL = CreateObject("Excel.Application")
    If Err = 0 Then

        ' Lanza siempre estea la nueva instanza Excel.
        On Error GoTo AbreExcelERR

        Set oXL = CreateObject("Excel.application")
    
        '    oXL.Visible = True: oXL.UserControl = True
        Set objWorkBook = Excel.Workbooks.Add
    
        '    Set objWorkBook = Excel.Workbooks.Add
        '     oXL.Visible = True: oXL.UserControl = True
    End If

    On Error GoTo AbreExcelERR

    oXL.Visible = True

    AbreExcel = True
    Exit Function
AbreExcelERR:
    AbreExcel = False
    Debug.Print "AbreExcel", Err, error
    MsgBox "gab n.:" & Str(Err) & " = " & error, , "excel.bas:apriExcel"
    MsgBox "Non è possibile lanciare in esecuzione Microsoft Excel.", MB_ICONSTOP
    Exit Function

End Function

Public Sub carga_imagen()
    Range("A1").Select

    Dim picname As String

    On Error GoTo carga_imagen

    picname = Range("H1")
    'ActiveSheet.Pictures.Insert("C:\Users\Administrador\Desktop\logo.jpg").Select
    'inicio
    'FileName = globaldir & "\temporal\" & gusuario & ".txt"
    FileName = globalpath & "\ico\" & "logo.jpg"
    ActiveSheet.Pictures.Insert(FileName).Select

    With Selection
        .Left$ = Range("A1").Left
        .Top = Range("A1").Top
        .ShapeRange.LockAspectRatio = msofalse

        '.ShapeRange.Height = 100#
        '.ShapeRange.Width = 80#
        '.ShapeRange.Rotation = 0#
    End With

    Range("A1").Select
    Application.ScreenUpdating = True
    Exit Sub
carga_imagen:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

End Sub

'Public Function titulo_solo_documentos(objWorkBook As Excel.Workbook, _
'                my_moneda As String)
                
Public Function titulo_solo_documentos(my_moneda As String)

    Set objWorksheet = objWorkBook.Worksheets(1)

    On Error GoTo titulo_solo_documentos

    objWorksheet.Cells(11, 8).Font.bold = True
    objWorksheet.Cells(11, 8) = "Moneda"
    objWorksheet.Cells(11, 9) = my_moneda
  
    'fecha a imprimir

    objWorksheet.Cells(11, 6).Font.bold = True
    'objWorksheet.Range("F11").ColumnWidth = 40
    objWorksheet.Cells(11, 6) = "FECHA HOY :"
    objWorksheet.Cells(11, 7) = Format(Now, "dd/mm/yyyy") & "- HORA " + Format(Now, "HH:MM:SS")
 
    'Aqui los margenes cuadrados
    objWorksheet.Range(Cells(12, 1), Cells(12, 9)).Borders.LineStyle = xlContinuous 'Codigo
    objWorksheet.Range(Cells(12, 1), Cells(12, 9)).Interior.color = RGB(215, 215, 215) 'Codigo
    objWorksheet.Range(Cells(12, 1), Cells(12, 9)).Font.bold = True
    objWorksheet.Range(Cells(12, 1), Cells(12, 9)).Font.Size = 14
  
    objWorksheet.Cells(12, 1).Font.bold = True
    objWorksheet.Cells(12, 1) = "Codigo"
    objWorksheet.Range("A12").ColumnWidth = 12
  
    objWorksheet.Cells(12, 2).Font.bold = True
    objWorksheet.Cells(12, 2) = "Nombre"
    objWorksheet.Range("B12").ColumnWidth = 30
  
    objWorksheet.Cells(12, 3).Font.bold = True
    objWorksheet.Cells(12, 3) = "Local"
    objWorksheet.Range("C12").ColumnWidth = 8
  
    objWorksheet.Cells(12, 4).Font.bold = True
    objWorksheet.Cells(12, 4) = "Estado"
    objWorksheet.Range("D12").ColumnWidth = 10
  
    objWorksheet.Cells(12, 5).Font.bold = True
    objWorksheet.Cells(12, 5) = "Tipo"
    objWorksheet.Range("E12").ColumnWidth = 8
  
    objWorksheet.Cells(12, 6).Font.bold = True
    objWorksheet.Cells(12, 6) = "Serie"
    objWorksheet.Range("F12").ColumnWidth = 10
  
    objWorksheet.Cells(12, 7).Font.bold = True
    objWorksheet.Cells(12, 7) = "Numero"
    objWorksheet.Range("G12").ColumnWidth = 10
  
    objWorksheet.Cells(12, 8).Font.bold = True
    objWorksheet.Cells(12, 8) = "Fecha"
    objWorksheet.Range("H12").ColumnWidth = 10

    objWorksheet.Cells(12, 9).Font.bold = True
    objWorksheet.Cells(12, 9) = "Total"
    objWorksheet.Range("I12").ColumnWidth = 8
    Exit Function
titulo_solo_documentos:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

End Function

Public Function titulo_documentos_sele(my_transporte As String)

    Dim c As Integer

    On Error GoTo titulo_documentos_sele

    'Set objWorksheet = objWorkBook.Worksheets(1)

    ' If my_transporte = "" Then
    '    c = 12
    ' Else
    c = 16
    ' End If
    'Aqui los margenes cuadrados
    '  objWorksheet.Range(Cells(12, 1), Cells(12, 7)).Borders.LineStyle = xlContinuous 'Codigo
    '  objWorksheet.Range(Cells(12, 1), Cells(12, 7)).Interior.color = RGB(215, 215, 215) 'Codigo
    '
    '  objWorksheet.Cells(12, 1).Font.bold = True
    '  objWorksheet.Cells(12, 1) = "Producto"
    '  objWorksheet.Range("A12").ColumnWidth = 8
    '
    '  objWorksheet.Cells(12, 2).Font.bold = True
    '  objWorksheet.Cells(12, 2) = "Descripcion"
    '  objWorksheet.Range("B12").ColumnWidth = 35
    '
    '  objWorksheet.Cells(12, 3).Font.bold = True
    '  objWorksheet.Cells(12, 3) = "Unidad"
    '  objWorksheet.Range("C12").ColumnWidth = 6
    '
    '  objWorksheet.Cells(12, 4).Font.bold = True
    '  objWorksheet.Cells(12, 4) = "Factor"
    '  'objWorksheet.Range("D12").ColumnWidth = 6
    '
    '  objWorksheet.Cells(12, 5).Font.bold = True
    '  objWorksheet.Cells(12, 5) = "Cantidad"
    '  objWorksheet.Range("E12").ColumnWidth = 12
    '
    '  objWorksheet.Cells(12, 6).Font.bold = True
    '  objWorksheet.Cells(12, 6) = "Precio"
    '  objWorksheet.Range("F12").ColumnWidth = 12
    '
    '  objWorksheet.Cells(12, 7).Font.bold = True
    '  objWorksheet.Cells(12, 7) = "Total"
    '  objWorksheet.Range("G12").ColumnWidth = 8
    '''
    objWorksheet.Range(Cells(c, 1), Cells(c, 7)).Borders.LineStyle = xlContinuous 'Codigo
    objWorksheet.Range(Cells(c, 1), Cells(c, 7)).Interior.color = RGB(215, 215, 215) 'Codigo

    objWorksheet.Cells(c, 1).Font.bold = True
    objWorksheet.Cells(c, 1) = "Producto"
    objWorksheet.Range("A12").ColumnWidth = 8
  
    objWorksheet.Cells(c, 2).Font.bold = True
    objWorksheet.Cells(c, 2) = "Descripcion"
    objWorksheet.Range("B12").ColumnWidth = 50
  
    objWorksheet.Cells(c, 3).Font.bold = True
    objWorksheet.Cells(c, 3) = "Unidad"
    objWorksheet.Range("C12").ColumnWidth = 6
  
    objWorksheet.Cells(c, 4).Font.bold = True
    objWorksheet.Cells(c, 4) = "FTR"
    objWorksheet.Range("D12").ColumnWidth = 4
  
    objWorksheet.Cells(c, 5).Font.bold = True
    objWorksheet.Cells(c, 5) = "Cantidad"
    objWorksheet.Range("E12").ColumnWidth = 12
  
    objWorksheet.Cells(c, 6).Font.bold = True
    objWorksheet.Cells(c, 6) = "Precio"
    objWorksheet.Range("F12").ColumnWidth = 12
  
    objWorksheet.Cells(c, 7).Font.bold = True
    objWorksheet.Cells(c, 7) = "Total"
    objWorksheet.Range("G12").ColumnWidth = 8
    Exit Function
titulo_documentos_sele:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            ' MsgBox Err.Number & vbcrlf & Err.Description
    End Select

End Function

Public Sub datos_documentos_sele(objWorkBook As Excel.Workbook, _
                                 v As Integer, _
                                 h As Integer, _
                                 sdx1 As Integer, _
                                 sdx2 As Integer)

    Set objWorksheet = objWorkBook.Worksheets(1)

    v = 13
    h = 0
    sdx = 0
    sdx1 = 0
    sdx2 = 0

    If rexplorap.RecordCount > 0 Then  'si existe

        Do

            If rexplorap.EOF Then Exit Do
            objWorksheet.Cells(v, h + 1) = rexplorap.Fields("Codigo")
            objWorksheet.Cells(v, h + 2) = rexplorap.Fields("nombre")
            objWorksheet.Cells(v, h + 3) = rexplorap.Fields("Local")
     
            If rexplorap.Fields("estado") = 2 Then
                objWorksheet.Cells(v, h + 4) = "Cerrado"
            ElseIf rexplorap.Fields("estado") = 0 Then
                objWorksheet.Cells(v, h + 4) = "Modifica"
            ElseIf rexplorap.Fields("estado") = 1 Then
                objWorksheet.Cells(v, h + 4) = "Anulado"

            End If
     
            objWorksheet.Cells(v, h + 5) = rexplorap.Fields("tipo")
            objWorksheet.Cells(v, h + 6) = rexplorap.Fields("serie")
            objWorksheet.Cells(v, h + 7) = rexplorap.Fields("numero")
            objWorksheet.Cells(v, h + 8) = rexplorap.Fields("Moneda")
            objWorksheet.Cells(v, h + 9) = rexplorap.Fields("fecha")
            objWorksheet.Cells(v, h + 10) = rexplorap.Fields("total")

            If "" & rexplorap.Fields("Moneda") = "S" Then
                sdx1 = sdx1 + Val("" & rexplorap.Fields("total"))

            End If

            If "" & rexplorap.Fields("Moneda") = "D" Then
                sdx2 = sdx2 + Val("" & rexplorap.Fields("total"))

            End If

            v = v + 1
            rexplorap.MoveNext
   
        Loop

    End If

End Sub

Public Sub detalle_documentos_sele(my_local As String, _
                                   my_tipo As String, _
                                   my_serie As String, _
                                   my_numero As String, _
                                   my_struc_cotizacion_total_excel() As struc_cotizacion_total_excel, _
                                   salida As Boolean, _
                                   D As Integer, _
                                   my_transporte As String)
           
    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    ReDim my_struc_cotizacion_total_excel(0)

    ' mysql = "SELECT distinct d.producto,d.descripcio,d.unidad,d.factor, " & Chr$(10)
    ' mysql = mysql & "d.cantidad,d.precio,d.total,d.subtotal" & Chr$(10)
    ' mysql = mysql & "From " & dgusuariog & " d," & Chr$(10)
    ' mysql = mysql & "factura f" & Chr$(10)
    ' mysql = mysql & "where d.local='" & Trim("" & my_local) & "' " & Chr$(10)
    ' mysql = mysql & "and d.tipo='" & Trim("" & my_tipo) & "' " & Chr$(10)
    ' mysql = mysql & "and d.serie='" & Trim("" & my_serie) & "' " & Chr$(10)
    ' mysql = mysql & "and d.numero='" & Trim("" & my_numero) & "'" & Chr$(10)
    ' mysql = mysql & "and d.tipo= f.TIPO" & Chr$(10)
    ' mysql = mysql & "and d.serie = f.SERIE" & Chr$(10)
    ' mysql = mysql & "and d.NUMERO = f.NUMERO" & Chr$(10)
    mysql = "SELECT * from " & dgusuariog & Chr$(10)
    mysql = mysql & "where local='" & Trim("" & my_local) & "' " & Chr$(10)
    mysql = mysql & "and tipo='" & Trim("" & my_tipo) & "' " & Chr$(10)
    mysql = mysql & "and serie='" & Trim("" & my_serie) & "' " & Chr$(10)
    mysql = mysql & "and numero='" & Trim("" & my_numero) & "'" & Chr$(10)
 
    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF Then
        salida = False
        Exit Sub
    Else
        salida = True
        mytablex.MoveFirst
        D = 0

        Do Until mytablex.EOF

            If D > 0 Then
                ReDim Preserve my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel) + 1)

            End If
     
            If mytablex.Fields("producto") <> "" Then
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).producto = mytablex.Fields("producto")
            Else
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).producto = ""

            End If
      
            If mytablex.Fields("descripcio") <> "" Then
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).descripcion = mytablex.Fields("descripcio")
            Else
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).descripcion = ""

            End If
    
            If mytablex.Fields("unidad") <> "" Then
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).unidad = mytablex.Fields("unidad")
            Else
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).unidad = ""

            End If
    
            If mytablex.Fields("factor") <> "" Then
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).factor = mytablex.Fields("factor")
            Else
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).factor = ""

            End If
   
            If mytablex.Fields("cantidad") <> "" Then
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).cantidad = mytablex.Fields("cantidad")
            Else
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).cantidad = 0

            End If
   
            If mytablex.Fields("precio") <> "" Then
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).precio = mytablex.Fields("precio")
            Else
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).precio = 0

            End If
   
            If mytablex.Fields("total") <> "" Then
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).total = mytablex.Fields("total")
            Else
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).total = 0

            End If

            If mytablex.Fields("subtotal") <> "" Then
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).subtotal = mytablex.Fields("subtotal")
            Else
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).subtotal = 0

            End If
   
            If mytablex.Fields("impuesto") <> "" Then
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).igv = mytablex.Fields("impuesto")
            Else
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).igv = 0

            End If
   
            D = D + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
    Exit Sub

End Sub

Public Sub titulo_familia_producto(fechai As String, _
                                   my_moneda As String, _
                                   fechaf As String)
           
    On Error GoTo titulo_familia_producto

    Set objWorksheet = objWorkBook.Worksheets(1)

    objWorksheet.Cells(9, 1).Font.bold = True
    objWorksheet.Cells(9, 1) = "Desde"
    objWorksheet.Cells(9, 1).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(9, 2).Select
    Selection.ColumnWidth = 9
    objWorksheet.Cells(9, 2).Font.Size = 4 'aqui tamaño letra
    'objWorksheet.Range("B9").ColumnWidth = 12
    objWorksheet.Cells(9, 2) = fechai
 
    objWorksheet.Cells(9, 3).Font.bold = True
    objWorksheet.Cells(9, 3).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(9, 3) = "Al"
    objWorksheet.Cells(9, 4).Select
    Selection.ColumnWidth = 9
    objWorksheet.Cells(9, 4).Font.Size = 4 'aqui tamaño letra
    objWorksheet.Cells(9, 4) = "'" & fechaf
    'moneda
    objWorksheet.Cells(9, 9).Font.bold = True
    objWorksheet.Cells(9, 9).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(9, 9) = "Moneda"
    objWorksheet.Cells(9, 10) = my_moneda
    'aqui pone la de la impresion
    objWorksheet.Cells(9, 7).Font.bold = True
    objWorksheet.Cells(9, 7).Font.Size = 4 'aqui tamaño letra
    objWorksheet.Cells(9, 7) = "FECHA HOY :"
    objWorksheet.Cells(9, 8).Font.Size = 4 'aqui tamaño letra
    objWorksheet.Cells(9, 8) = Format(Now, "dd/mm/yyyy") & "- HORA " + Format(Now, "HH:MM:SS")
 
    'Aqui los margenes cuadrados
    'objWorksheet.Range(Cells(10, 1), Cells(10, 11)).Borders.LineStyle = xlContinuous 'Codigo
    'objWorksheet.Range(Cells(10, 1), Cells(10, 11)).Interior.color = RGB(215, 215, 215) 'Codigo

    objWorksheet.Range(Cells(10, 1), Cells(10, 9)).Borders.LineStyle = xlContinuous 'Codigo
    objWorksheet.Range(Cells(10, 1), Cells(10, 9)).Interior.color = RGB(215, 215, 215) 'Codigo
  
    objWorksheet.Cells(10, 1).Font.bold = True
    objWorksheet.Cells(10, 1).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 1) = "Familia"
    objWorksheet.Range("A10").ColumnWidth = 7

    objWorksheet.Cells(10, 2).Font.bold = True
    objWorksheet.Cells(10, 2).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 2) = "Producto"
    objWorksheet.Range("B10").ColumnWidth = 6

    objWorksheet.Cells(10, 3).Font.bold = True
    objWorksheet.Cells(10, 3).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 3) = "Descripcion"
    objWorksheet.Range("C10").ColumnWidth = 25
 
    objWorksheet.Cells(10, 4).Font.bold = True
    objWorksheet.Cells(10, 4).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 4) = "Unid."
    objWorksheet.Range("D10").ColumnWidth = 5

    objWorksheet.Cells(10, 5).Font.bold = True
    objWorksheet.Cells(10, 5).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 5) = "Factor"
    objWorksheet.Range("E10").ColumnWidth = 4

    objWorksheet.Cells(10, 6).Font.bold = True
    objWorksheet.Cells(10, 6).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 6) = "Cantidad"
    objWorksheet.Range("F10").ColumnWidth = 6

    objWorksheet.Cells(10, 7).Font.bold = True
    objWorksheet.Cells(10, 7).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 7) = "Total"
    objWorksheet.Range("G10").ColumnWidth = 8

    objWorksheet.Cells(10, 8).Font.bold = True
    objWorksheet.Cells(10, 8).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 8) = "TCosto"
    objWorksheet.Range("H10").ColumnWidth = 12

    objWorksheet.Cells(10, 9).Font.bold = True
    objWorksheet.Cells(10, 9).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 9) = "Ganancia"
    objWorksheet.Range("I10").ColumnWidth = 8

    Exit Sub
titulo_familia_producto:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

End Sub

Public Sub suma_cmb_docu_sele(objWorkBook As Excel.Workbook)
    objExcel.ActiveSheet.Cells(v, h + 4) = "" & Format(suma1, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 5) = "" & Format(suma2, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 6) = "" & Format(suma5, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 7) = "" & Format(suma6, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 8) = "" & Format(suma3, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 9) = "" & Format(suma4, "0.00")

End Sub

Public Sub producto_docu_sele(v As Integer, _
                              h As Integer, _
                              my_struc_producto() As struc_producto, _
                              k As Integer)
          
    Dim mysql    As String

    Dim mytabley As New ADODB.Recordset

    v = 11
    h = 0
    sdx = 0
    sdx1 = 0
    sdx2 = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0
    suma6 = 0

    On Error GoTo producto_docu_sele

    Set objWorksheet = objWorkBook.Worksheets(1)

    For j = 0 To k - 1
        mysql = "SELECT pr.FAMILIA as familia,p.UNIDAD1,p.UNIDAD2,p.UNIDAD3," & Chr$(10)
        mysql = mysql & "p.unidad4 , p.unidad5, " & Chr$(10)
        mysql = mysql & "p.UNIDAD6,p.UNIDAD7,p.UNIDAD7,p.UNIDAD8,p.UNIDAD9,p.UNIDAD10," & Chr$(10)
        mysql = mysql & "p.factor1,p.factor2,p.factor3,p.factor4,p.factor5," & Chr$(10)
        mysql = mysql & "p.factor6,p.factor7,p.factor8,p.factor9,p.factor10," & Chr$(10)
        mysql = mysql & "p.pventa1,p.pventa2,p.pventa3,p.pventa4,p.pventa5," & Chr$(10)
        mysql = mysql & "p.pventa6,p.pventa7,p.pventa8,p.pventa9,p.pventa10," & Chr$(10)
        mysql = mysql & "max(d.fecha),pr.COSTOU as COSTOU " & Chr$(10)
        mysql = mysql & "from precios p," & Chr$(10)
        mysql = mysql & "producto pr," & Chr$(10)
        mysql = mysql & "detalle d" & Chr$(10)
        mysql = mysql & "where p.producto='" & "" & my_struc_producto(j).producto & "'" & Chr$(10)
        mysql = mysql & "AND pr.producto = p.producto" & Chr$(10)
        mysql = mysql & "and p.PRODUCTO = d.PRODUCTO" & Chr$(10)
        mysql = mysql & "Group By pr.FAMILIA," & Chr$(10)
        mysql = mysql & "p.UNIDAD1,p.UNIDAD2,p.UNIDAD3," & Chr$(10)
        mysql = mysql & "p.UNIDAD4,p.UNIDAD5,p.UNIDAD6," & Chr$(10)
        mysql = mysql & "p.UNIDAD7,p.UNIDAD8,p.UNIDAD9," & Chr$(10)
        mysql = mysql & "p.UNIDAD10," & Chr$(10)
        mysql = mysql & "p.FACTOR1,p.FACTOR2,p.FACTOR3," & Chr$(10)
        mysql = mysql & "p.FACTOR4,p.FACTOR5,p.FACTOR6," & Chr$(10)
        mysql = mysql & "p.FACTOR7,p.FACTOR8,p.FACTOR9,p.FACTOR10," & Chr$(10)
        mysql = mysql & "p.PVENTA1,p.PVENTA2,p.PVENTA3,p.PVENTA4," & Chr$(10)
        mysql = mysql & "p.PVENTA5,p.PVENTA6,p.PVENTA7,p.PVENTA8," & Chr$(10)
        mysql = mysql & "p.pventa9 , p.pventa10, pr.COSTOU" & Chr$(10)
  
        mytabley.Open mysql, cn, adOpenStatic, adLockOptimistic
 
        If mytabley.RecordCount > 0 Then  'si existe
            If mytabley.Fields("familia") <> my_familia Then
                objWorksheet.Cells(v, h + 1).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 1) = mytabley.Fields("familia")
                my_familia = mytabley.Fields("familia")

            End If

            objWorksheet.Cells(v, h + 2).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 2) = my_struc_producto(j).producto
    
            objWorksheet.Cells(v, h + 3).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 3) = my_struc_producto(j).descripcion

            If mytabley.Fields("unidad1") <> "" Then
                objWorksheet.Cells(v, h + 4).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 4) = "" & mytabley.Fields("unidad1")
            ElseIf mytabley.Fields("unidad2") <> "" Then
                objWorksheet.Cells(v, h + 4).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 4) = "" & mytabley.Fields("unidad2")
            ElseIf mytabley.Fields("unidad3") <> "" Then
                objWorksheet.Cells(v, h + 4).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 4) = "" & mytabley.Fields("unidad3")
            ElseIf mytabley.Fields("unidad4") <> "" Then
                objWorksheet.Cells(v, h + 4).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 4) = "" & mytabley.Fields("unidad4")
            ElseIf mytabley.Fields("unidad5") <> "" Then
                objWorksheet.Cells(v, h + 5).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 4) = "" & mytabley.Fields("unidad5")
            ElseIf mytabley.Fields("unidad6") <> "" Then
                objWorksheet.Cells(v, h + 4).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 4) = "" & mytabley.Fields("unidad6")
            ElseIf mytabley.Fields("unidad7") <> "" Then
                objWorksheet.Cells(v, h + 4).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 4) = "" & mytabley.Fields("unidad7")
            ElseIf mytabley.Fields("unidad8") <> "" Then
                objWorksheet.Cells(v, h + 4).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 4) = "" & mytabley.Fields("unidad8")
            ElseIf mytabley.Fields("unidad9") <> "" Then
                objWorksheet.Cells(v, h + 4).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 4) = "" & mytabley.Fields("unidad9")
            ElseIf mytabley.Fields("unidad10") <> "" Then
                objWorksheet.Cells(v, h + 4).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 4) = "" & mytabley.Fields("unidad10")

            End If
   
            If mytabley.Fields("factor1") <> "" Then
                objWorksheet.Cells(v, h + 5).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 5) = "" & mytabley.Fields("factor1")
                xfactor = "" & mytabley.Fields("factor1")
            ElseIf mytabley.Fields("factor2") <> "" Then
                objWorksheet.Cells(v, h + 5).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 5) = "" & mytabley.Fields("factor2")
                xfactor = "" & mytabley.Fields("factor2")
            ElseIf mytabley.Fields("factor3") <> "" Then
                objWorksheet.Cells(v, h + 5).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 5) = "" & mytabley.Fields("factor3")
                xfactor = "" & mytabley.Fields("factor3")
            ElseIf mytabley.Fields("factor4") <> "" Then
                objWorksheet.Cells(v, h + 5).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 5) = "" & mytabley.Fields("factor4")
                xfactor = "" & mytabley.Fields("factor4")
            ElseIf mytabley.Fields("factor5") <> "" Then
                objWorksheet.Cells(v, h + 5).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 5) = "" & mytabley.Fields("factor5")
                xfactor = "" & mytabley.Fields("factor5")
            ElseIf mytabley.Fields("factor6") <> "" Then
                objWorksheet.Cells(v, h + 5).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 5) = "" & mytabley.Fields("factor6")
                xfactor = "" & mytabley.Fields("factor6")
            ElseIf mytabley.Fields("factor7") <> "" Then
                objWorksheet.Cells(v, h + 5).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 5) = "" & mytabley.Fields("factor7")
                xfactor = "" & mytabley.Fields("factor7")
            ElseIf mytabley.Fields("factor8") <> "" Then
                objWorksheet.Cells(v, h + 5).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 5) = "" & mytabley.Fields("factor8")
                xfactor = "" & mytabley.Fields("factor8")
            ElseIf mytabley.Fields("factor9") <> "" Then
                objWorksheet.Cells(v, h + 5).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 5) = "" & mytabley.Fields("factor9")
                xfactor = "" & mytabley.Fields("factor9")
            ElseIf mytabley.Fields("factor10") <> "" Then
                objWorksheet.Cells(v, h + 5).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 5) = "" & mytabley.Fields("factor10")
                xfactor = "" & mytabley.Fields("factor10")

            End If

            buf = calcula_saldo(Val("" & my_struc_producto(j).xcanti), Val(xfactor))
            objWorksheet.Cells(v, h + 6).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 6) = buf 'CANTIDAD
            'precio unitario

            If mytabley.Fields("pventa1") <> "" Then
                objWorksheet.Cells(v, h + 7).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 7) = "" & mytabley.Fields("pventa1")
                Punitario = "" & mytabley.Fields("pventa1")
            ElseIf mytabley.Fields("pventa2") <> "" Then
                objWorksheet.Cells(v, h + 7).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 7) = "" & mytabley.Fields("pventa2")
                Punitario = "" & mytabley.Fields("pventa2")
            ElseIf mytabley.Fields("pventa3") <> "" Then
                objWorksheet.Cells(v, h + 7).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 7) = "" & mytabley.Fields("pventa3")
                Punitario = "" & mytabley.Fields("pventa3")
            ElseIf mytabley.Fields("pventa4") <> "" Then
                objWorksheet.Cells(v, h + 7).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 7) = "" & mytabley.Fields("pventa4")
                Punitario = "" & mytabley.Fields("pventa4")
            ElseIf mytabley.Fields("pventa5") <> "" Then
                objWorksheet.Cells(v, h + 7).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 7) = "" & mytabley.Fields("pventa5")
                Punitario = "" & mytabley.Fields("pventa5")
            ElseIf mytabley.Fields("pventa6") <> "" Then
                objWorksheet.Cells(v, h + 7).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 7) = "" & mytabley.Fields("pventa6")
                Punitario = "" & mytabley.Fields("pventa6")
            ElseIf mytabley.Fields("pventa7") <> "" Then
                objWorksheet.Cells(v, h + 7).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 7) = "" & mytabley.Fields("pventa7")
                Punitario = "" & mytabley.Fields("pventa7")
            ElseIf mytabley.Fields("pventa8") <> "" Then
                objWorksheet.Cells(v, h + 7).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 7) = "" & mytabley.Fields("pventa8")
                Punitario = "" & mytabley.Fields("pventa8")
            ElseIf mytabley.Fields("pventa9") <> "" Then
                objWorksheet.Cells(v, h + 7).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 7) = "" & mytabley.Fields("pventa9")
                Punitario = "" & mytabley.Fields("pventa9")
            ElseIf mytabley.Fields("pventa10") <> "" Then
                objWorksheet.Cells(v, h + 7).Font.Size = 7 'aqui tamaño letra
                objWorksheet.Cells(v, h + 7) = "" & mytabley.Fields("pventa10")
                Punitario = "" & mytabley.Fields("pventa10")

            End If

            objWorksheet.Cells(v, h + 8).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 8) = Punitario * buf    'vtaxproc.total
            objWorksheet.Cells(v, h + 9).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 9) = "" & mytabley.Fields("COSTOU") 'costo unitario
            'para los calculos
            sdx = Val(ccosto) / Val(xfactor)
            'inicio 04/07/2017
            'sdx1 = Val("" & my_struc_producto(j).xtotal - Val(ccosto) * Val("" & my_struc_producto(j).xcanti))
            sdx1 = Val("" & my_struc_producto(j).xtotal - mytabley.Fields("COSTOU"))
   
            objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 10) = "" & Format(my_struc_producto(j).xtotal - Val(xcosto) * my_struc_producto(j).xcanti, , "0.00")  'costo total
            objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 11) = sdx1 'ganancia
                                  
            v = v + 1
            mytabley.MoveNext

            '  'aqui  acumuladores
            If my_struc_producto(j).moneda = "S" Then
                suma1 = suma1 + Val("" & my_struc_producto(j).xcanti)
                suma2 = suma2 + Val("" & my_struc_producto(j).xtotal)
                'suma4 = suma4 + Val("" & my_struc_producto(j).xtotal - Val(ccosto) * Val("" & my_struc_producto(j).xcanti)) 'ganancia
                suma4 = suma4 + Val("" & my_struc_producto(j).xtotal - mytabley.Fields("COSTOU"))
                suma3 = suma3 + Punitario * buf 'vtaxproc.total
    
                suma5 = suma5 + sdx 'costo unitario
                suma6 = suma6 + sdx1 'costo total

            End If

            If my_struc_producto(j).moneda = "D" Then
                suma3 = suma3 + Val("" & my_struc_producto(j).xcanti)
                'suma4 = suma4 + Val("" & my_struc_producto(j).xtotal)
                suma4 = suma4 + sdx1

            End If

        End If

        mytabley.Close
    Next j

    'aqui los sub-totales
    objWorksheet.Cells(v, h + 5).Font.bold = True
    objWorksheet.Cells(v, h + 5).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 5) = "Sub-TOTALES"
   
    objWorksheet.Cells(v, h + 6).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 6) = "" & Format(suma1, "0.00") 'cantidad
   
    'objWorksheet.Cells(v, h + 7).Interior.color = RGB(215, 215, 0)  'resaltador
    'objWorksheet.Cells(v, h + 7) = "" & Format(suma2, "0.00") 'precio unitario
   
    objWorksheet.Cells(v, h + 6).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 6) = "" & Format(suma3, "0.00") ''vtaxproc.total
   
    objWorksheet.Cells(v, h + 9).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 9) = "" & Format(suma5, "0.00") 'costo unitario
   
    objWorksheet.Cells(v, h + 10).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 10) = "" & Format(suma6, "0.00") 'costo total
   
    objWorksheet.Cells(v, h + 11).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 11) = "" & Format(suma4, "0.00") 'ganancia
    v = v + 1
    'aqui los totales
    'provemos anteriormente tambien estaba asi
    objWorksheet.Cells(v, h + 5).Font.bold = True
    objWorksheet.Cells(v, h + 5).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 5) = "TOTALES"
   
    objWorksheet.Cells(v, h + 6).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 6) = "" & Format(suma1, "0.00") 'cantidad
   
    'objWorksheet.Cells(v, h + 7).Interior.color = RGB(215, 215, 0)  'resaltador
    'objWorksheet.Cells(v, h + 7) = "" & Format(suma2, "0.00") 'precio unitario
   
    objWorksheet.Cells(v, h + 6).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 6) = "" & Format(suma3, "0.00") ''vtaxproc.total
   
    objWorksheet.Cells(v, h + 9).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 9) = "" & Format(suma5, "0.00") 'costo unitario
   
    objWorksheet.Cells(v, h + 10).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 10) = "" & Format(suma6, "0.00") 'costo total
   
    objWorksheet.Cells(v, h + 11).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 11) = "" & Format(suma4, "0.00") 'ganancia
    Exit Sub
producto_docu_sele:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

End Sub

Public Sub Carga_fproducto(v As Integer, _
                           h As Integer, _
                           my_struc_producto() As struc_producto, _
                           k As Integer)
          
    Dim mysql    As String

    Dim mytabley As New ADODB.Recordset

    v = 11
    h = 0
    sdx = 0
    sdx1 = 0
    sdx2 = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0
    suma6 = 0

    On Error GoTo producto_docu_sele

    Set objWorksheet = objWorkBook.Worksheets(1)

    For j = 0 To k - 1
        objWorksheet.Cells(v, h + 1).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 1) = my_struc_producto(j).familia
   
        objWorksheet.Cells(v, h + 2).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 2) = my_struc_producto(j).producto
    
        objWorksheet.Cells(v, h + 3).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 3) = my_struc_producto(j).descripcion
    
        objWorksheet.Cells(v, h + 4).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 4) = "" & my_struc_producto(j).unidad
    
        objWorksheet.Cells(v, h + 5).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 5) = "" & my_struc_producto(j).factor
    
        buf = calcula_saldo(Val("" & my_struc_producto(j).xcanti), Val(my_struc_producto(j).factor))
        objWorksheet.Cells(v, h + 6).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 6) = buf 'CANTIDAD
  
        'total
        objWorksheet.Cells(v, h + 7).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 7) = Format(my_struc_producto(j).xtotal)
  
        'tcosto
        objWorksheet.Cells(v, h + 8).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 8) = Val(my_struc_producto(j).xcostou * my_struc_producto(j).xcanti)
  
        'ganancia
        sdx = Val(Format(my_struc_producto(j).xcostou / Val(my_struc_producto(j).factor)))
        objWorksheet.Cells(v, h + 9).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 9) = sdx 'ganancia
                                  
        v = v + 1

        '  'aqui  acumuladores
        If my_struc_producto(j).moneda = "S" Or my_struc_producto(j).moneda = "D" Then
            buf = calcula_saldo(Val("" & my_struc_producto(j).xcanti), Val(my_struc_producto(j).factor))
            suma1 = suma1 + buf
    
            suma2 = suma2 + Val("" & my_struc_producto(j).xtotal)
            'suma4 = suma4 + Val("" & my_struc_producto(j).xtotal - Val(ccosto) * Val("" & my_struc_producto(j).xcanti)) 'ganancia
            sdx = Val(Format(my_struc_producto(j).xcostou / Val(my_struc_producto(j).factor)))
            suma4 = suma4 + sdx 'ganancia
            suma3 = suma3 + Punitario * buf 'vtaxproc.total
    
            suma5 = suma5 + sdx 'costo unitario
            suma6 = suma6 + sdx1 'costo total

        End If
  
    Next j

    'aqui los sub-totales
    objWorksheet.Cells(v, h + 5).Font.bold = True
    objWorksheet.Cells(v, h + 5).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 5) = "Sub-TOTALES"
   
    objWorksheet.Cells(v, h + 6).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 6) = "" & Format(suma1, "0.00") 'cantidad
   
    objWorksheet.Cells(v, h + 7).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 7) = "" & Format(suma2, "0.00") 'precio unitario
   
    'objWorksheet.Cells(v, h + 6).Interior.color = RGB(215, 215, 0)  'resaltador
    ' objWorksheet.Cells(v, h + 6) = "" & Format(suma3, "0.00") ''vtaxproc.total
   
    objWorksheet.Cells(v, h + 9).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 9) = "" & Format(suma4, "0.00") 'ganancia
   
    'objWorksheet.Cells(v, h + 10).Interior.color = RGB(215, 215, 0)  'resaltador
    'objWorksheet.Cells(v, h + 10) = "" & Format(suma6, "0.00") 'costo total
   
    'objWorksheet.Cells(v, h + 11).Interior.color = RGB(215, 215, 0)  'resaltador
    'objWorksheet.Cells(v, h + 11) = "" & Format(suma4, "0.00") 'ganancia
    v = v + 1
    'aqui los totales
    'provemos anteriormente tambien estaba asi
    objWorksheet.Cells(v, h + 5).Font.bold = True
    objWorksheet.Cells(v, h + 5).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 5) = "TOTALES"
   
    objWorksheet.Cells(v, h + 6).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 6) = "" & Format(suma1, "0.00") 'cantidad
   
    'objWorksheet.Cells(v, h + 7).Interior.color = RGB(215, 215, 0)  'resaltador
    'objWorksheet.Cells(v, h + 7) = "" & Format(suma2, "0.00") 'precio unitario
   
    'objWorksheet.Cells(v, h + 6).Interior.color = RGB(215, 215, 0)  'resaltador
    'objWorksheet.Cells(v, h + 6) = "" & Format(suma3, "0.00") ''vtaxproc.total
   
    objWorksheet.Cells(v, h + 9).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 9) = "" & Format(suma4, "0.00") 'ganancia
   
    'objWorksheet.Cells(v, h + 10).Interior.color = RGB(215, 215, 0)  'resaltador
    'objWorksheet.Cells(v, h + 10) = "" & Format(suma6, "0.00") 'costo total
   
    'objWorksheet.Cells(v, h + 11).Interior.color = RGB(215, 215, 0)  'resaltador
    'objWorksheet.Cells(v, h + 11) = "" & Format(suma4, "0.00") 'ganancia
    Exit Sub
producto_docu_sele:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

End Sub

Public Sub total_docu_sele(mytablex As ADODB.Recordset, _
                           objWorkBook As Excel.Workbook, _
                           v As Integer, _
                           h As Integer, _
                           suma1 As Double, _
                           suma2 As Double, _
                           suma3 As Double, _
                           suma4 As Double, _
                           suma5 As Double, _
                           suma6 As Double)

    Set objWorksheet = objWorkBook.Worksheets(1)
   
    objWorksheet.Cells(v, h + 4).Font.bold = True
    objWorksheet.Cells(v, h + 4).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 4) = "Sub-TOTALES"
   
    objWorksheet.Cells(v, h + 5).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 5) = "" & Format(suma1, "0.00") 'cantidad
   
    objWorksheet.Cells(v, h + 6).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 6) = "" & Format(suma2, "0.00") 'precio unitario
   
    objWorksheet.Cells(v, h + 7).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 7) = "" & Format(suma3, "0.00") ''vtaxproc.total
   
    objWorksheet.Cells(v, h + 8).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 8) = "" & Format(suma5, "0.00") 'costo unitario
   
    objWorksheet.Cells(v, h + 9).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 9) = "" & Format(suma6, "0.00") 'costo total
   
    objWorksheet.Cells(v, h + 10).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 10) = "" & Format(suma4, "0.00") 'ganancia

    v = v + 1
    'provemos anteriormente tambien estaba asi
    objWorksheet.Cells(v, h + 4).Font.bold = True
    objWorksheet.Cells(v, h + 4).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 4) = "TOTALES"
   
    objWorksheet.Cells(v, h + 5).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 5) = "" & Format(suma1, "0.00") 'cantidad
   
    objWorksheet.Cells(v, h + 6).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 6) = "" & Format(suma2, "0.00") 'precio unitario
   
    objWorksheet.Cells(v, h + 7).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 7) = "" & Format(suma3, "0.00") ''vtaxproc.total
   
    objWorksheet.Cells(v, h + 8).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 8) = "" & Format(suma5, "0.00") 'costo unitario
   
    objWorksheet.Cells(v, h + 9).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 9) = "" & Format(suma6, "0.00") 'costo total
   
    objWorksheet.Cells(v, h + 10).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 10) = "" & Format(suma4, "0.00") 'ganancia

End Sub

Public Sub titulo_Registro_ventas(fechai As String, fechaf As String, acu As String)
           
    On Error GoTo titulo_Registro_ventas

    Set objWorksheet = objWorkBook.Worksheets(1)

    objWorksheet.Cells(9, 1).Font.bold = True
    objWorksheet.Cells(9, 1).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(9, 1) = "Desde"
    objWorksheet.Cells(9, 2).Select
    Selection.ColumnWidth = 9
    objWorksheet.Cells(9, 2).Font.Size = 4 'aqui tamaño letra
    objWorksheet.Cells(9, 2) = fechai
 
    objWorksheet.Cells(9, 4).Font.bold = True
    objWorksheet.Cells(9, 4).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(9, 4) = "Al"
    'objWorksheet.Range("E9").ColumnWidth = 9
    objWorksheet.Cells(9, 5).Select
    Selection.ColumnWidth = 9
    objWorksheet.Cells(9, 5).Font.Size = 4 'aqui tamaño letra
    objWorksheet.Cells(9, 5) = fechaf
 
    If acu = "V" Then
        objWorksheet.Cells(9, 8).Font.bold = True
        objWorksheet.Cells(9, 8) = "Registro Ventas"
        objWorksheet.Cells(9, 8).Font.Size = 7 'aqui tamaño letra
    Else
        objWorksheet.Cells(9, 8).Font.bold = True
        objWorksheet.Cells(9, 8) = "Registro Compras"
        objWorksheet.Cells(9, 8).Font.Size = 7 'aqui tamaño letra

    End If

    objWorksheet.Cells(9, 12).Font.bold = True
    objWorksheet.Cells(9, 12).Font.Size = 4 'aqui tamaño letra
    objWorksheet.Cells(9, 12) = "FECHA HOY :"
    objWorksheet.Cells(9, 13) = Format(Now, "dd/mm/yyyy") & "- HORA " + Format(Now, "HH:MM:SS")
    objWorksheet.Cells(9, 13).Font.Size = 4 'aqui tamaño letra
    'Aqui los margenes cuadrados
    objWorksheet.Range(Cells(10, 1), Cells(10, 17)).Borders.LineStyle = xlContinuous 'Codigo
    objWorksheet.Range(Cells(10, 1), Cells(10, 17)).Interior.color = RGB(215, 215, 215) 'Codigo

    objWorksheet.Cells(10, 1).Font.bold = True
    objWorksheet.Cells(10, 1).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 1) = "Fecha"
    objWorksheet.Range("A10").ColumnWidth = 8
  
    objWorksheet.Cells(10, 2).Font.bold = True
    objWorksheet.Cells(10, 2).Select
    Selection.ColumnWidth = 6
    objWorksheet.Cells(10, 2).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 2) = "Local"

    objWorksheet.Cells(10, 3).Font.bold = True
    objWorksheet.Cells(10, 3).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 3) = "Tipo"
    objWorksheet.Range("C10").ColumnWidth = 3
  
    objWorksheet.Cells(10, 4).Font.bold = True
    objWorksheet.Cells(10, 4).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 4) = "Serie"
    objWorksheet.Range("D10").ColumnWidth = 6

    objWorksheet.Cells(10, 5).Font.bold = True
    objWorksheet.Cells(10, 5).Select
    Selection.ColumnWidth = 6
    objWorksheet.Cells(10, 5).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 5) = "Numero"

    objWorksheet.Cells(10, 6).Font.bold = True
    objWorksheet.Cells(10, 6).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 6) = "Codigo"
    objWorksheet.Range("F10").ColumnWidth = 8
  
    objWorksheet.Cells(10, 7).Font.bold = True
    objWorksheet.Cells(10, 7).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 7) = "Clientes"
    objWorksheet.Range("G10").ColumnWidth = 10
  
    objWorksheet.Cells(10, 8).Font.bold = True
    objWorksheet.Cells(10, 8).Select
    Selection.ColumnWidth = 6
    objWorksheet.Cells(10, 8).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 8) = "Estado"
  
    objWorksheet.Cells(10, 9).Font.bold = True
    objWorksheet.Cells(10, 9).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 9) = "BaseImp."
    objWorksheet.Range("I10").ColumnWidth = 8 '4
   
    objWorksheet.Cells(10, 10).Font.bold = True
    objWorksheet.Cells(10, 10).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 10) = "Exra."
    objWorksheet.Range("J10").ColumnWidth = 4
  
    objWorksheet.Cells(10, 11).Font.bold = True
    objWorksheet.Cells(10, 11).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 11) = "Isc"
    objWorksheet.Range("K10").ColumnWidth = 4
  
    objWorksheet.Cells(10, 12).Font.bold = True
    objWorksheet.Cells(10, 12).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 12) = "I.G.V."
    objWorksheet.Range("L10").ColumnWidth = 6
    
    objWorksheet.Cells(10, 13).Font.bold = True
    objWorksheet.Cells(10, 13).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 13) = "Total"
    objWorksheet.Range("M10").ColumnWidth = 6
  
    objWorksheet.Cells(10, 14).Font.bold = True
    objWorksheet.Cells(10, 14).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 14) = "Ivap"
    objWorksheet.Range("N10").ColumnWidth = 6
  
    objWorksheet.Cells(10, 15).Font.bold = True
    objWorksheet.Cells(10, 15).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 15) = "Percepc."
    objWorksheet.Range("O10").ColumnWidth = 4
  
    objWorksheet.Cells(10, 16).Font.bold = True
    objWorksheet.Cells(10, 16).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 16) = "Serv."
    objWorksheet.Range("P10").ColumnWidth = 4
  
    objWorksheet.Cells(10, 17).Font.bold = True
    objWorksheet.Cells(10, 17).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 17) = "Detracc."
    objWorksheet.Range("Q10").ColumnWidth = 4
    Exit Sub
titulo_Registro_ventas:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

End Sub

Public Sub titulo_Movimiento_fpago(mytablex As ADODB.Recordset, _
                                   fechai As String, _
                                   fechaf As String)
           
    On Error GoTo titulo_Movimiento_fpago

    Set objWorksheet = objWorkBook.Worksheets(1)

    objWorksheet.Cells(9, 1).Font.bold = True
    objWorksheet.Cells(9, 1) = "Desde"
    objWorksheet.Cells(9, 2).Select
    Selection.ColumnWidth = 9
    objWorksheet.Cells(9, 2).Font.Size = 4 'aqui tamaño letra
    objWorksheet.Cells(9, 2) = fechai
 
    objWorksheet.Cells(9, 4).Font.bold = True
    objWorksheet.Cells(9, 4).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(9, 4) = "Al"
    objWorksheet.Cells(9, 5).Select
    Selection.ColumnWidth = 9
    objWorksheet.Cells(9, 5).Font.Size = 4 'aqui tamaño letra
    objWorksheet.Cells(9, 5) = fechaf

    objWorksheet.Cells(9, 11).Font.bold = True
    objWorksheet.Cells(9, 11).Font.Size = 4 'aqui tamaño letra
    objWorksheet.Cells(9, 11) = "FECHA HOY :"
    objWorksheet.Cells(9, 12) = Format(Now, "dd/mm/yyyy") & "- HORA " + Format(Now, "HH:MM:SS")
    'Aqui los margenes cuadrados
    objWorksheet.Range(Cells(10, 1), Cells(10, 14)).Borders.LineStyle = xlContinuous 'Codigo
    objWorksheet.Range(Cells(10, 1), Cells(10, 14)).Interior.color = RGB(215, 215, 215) 'Codigo

    objWorksheet.Cells(10, 1).Font.bold = True
    objWorksheet.Cells(10, 1).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 1) = "Tipo"
    objWorksheet.Range("A10").ColumnWidth = 8
  
    objWorksheet.Cells(10, 2).Font.bold = True
    objWorksheet.Cells(10, 2).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 2) = "Serie"
    objWorksheet.Range("B10").ColumnWidth = 6
  
    objWorksheet.Cells(10, 3).Font.bold = True
    objWorksheet.Cells(10, 3).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 3) = "Numero"
    objWorksheet.Range("C10").ColumnWidth = 6
  
    objWorksheet.Cells(10, 4).Font.bold = True
    objWorksheet.Cells(10, 4).Font.Size = 7 '
    objWorksheet.Cells(10, 4) = "Fecha"
    objWorksheet.Range("D10").ColumnWidth = 8
  
    objWorksheet.Cells(10, 5).Font.bold = True
    objWorksheet.Cells(10, 5).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 5) = "Total"
    objWorksheet.Range("E10").ColumnWidth = 8
  
    objWorksheet.Cells(10, 6).Font.bold = True
    objWorksheet.Cells(10, 6).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 6) = "Forma Pago"
    objWorksheet.Range("F10").ColumnWidth = 8
  
    objWorksheet.Cells(10, 7).Font.bold = True
    objWorksheet.Cells(10, 7).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 7) = "Descripcion"
    objWorksheet.Range("G10").ColumnWidth = 8
  
    objWorksheet.Cells(10, 8).Font.bold = True
    objWorksheet.Cells(10, 8).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 8) = "Orden"
    objWorksheet.Range("H10").ColumnWidth = 8
  
    objWorksheet.Cells(10, 9).Font.bold = True
    objWorksheet.Cells(10, 9).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 9) = "Estado"
    objWorksheet.Range("I10").ColumnWidth = 8
  
    objWorksheet.Cells(10, 10).Font.bold = True
    objWorksheet.Cells(10, 10).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 10) = "Codigo"
    objWorksheet.Range("J10").ColumnWidth = 8
  
    objWorksheet.Cells(10, 11).Font.bold = True
    objWorksheet.Cells(10, 11).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 11) = "Cliente"
    objWorksheet.Range("K10").ColumnWidth = 8
  
    objWorksheet.Cells(10, 12).Font.bold = True
    objWorksheet.Cells(10, 12).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 12) = "Usuario"
    objWorksheet.Range("L10").ColumnWidth = 6
  
    objWorksheet.Cells(10, 13).Font.bold = True
    objWorksheet.Cells(10, 13).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 13) = "Caja"
    objWorksheet.Range("M10").ColumnWidth = 6
  
    objWorksheet.Cells(10, 14).Font.bold = True
    objWorksheet.Cells(10, 14).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 14) = "Turno"
    objWorksheet.Range("010").ColumnWidth = 6
  
    objWorksheet.Cells(11, 8).Font.bold = True
    objWorksheet.Cells(11, 8) = "MONEDA"
    Exit Sub
titulo_Movimiento_fpago:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select
  
End Sub

Public Sub titulo_r_saldo_actual(fechai As String, fechaf As String, my_moneda As String)

    On Error GoTo titulo_r_saldo_actual

    Set objWorksheet = objWorkBook.Worksheets(1)

    If my_moneda = "Soles" Then
        objWorksheet.Cells(9, 9).Font.bold = True
        objWorksheet.Cells(9, 9).Font.Size = 5
        objWorksheet.Cells(9, 9).Select
        Selection.ColumnWidth = 9
                  
        objWorksheet.Cells(9, 9) = "Moneda"
        objWorksheet.Cells(9, 10).Font.Size = 5
        objWorksheet.Cells(9, 10) = "Soles"

        'objWorksheet.Cells(10, 11).Font.Size = 4
         
    Else
        objWorksheet.Cells(9, 9).Font.bold = True
        objWorksheet.Cells(9, 9).Font.Size = 5
        bjWorksheet.Cells(9, 9).Select
        Selection.ColumnWidth = 9
        objWorksheet.Cells(9, 9) = "Moneda"
        objWorksheet.Cells(9, 10).Font.Size = 5
        objWorksheet.Cells(9, 10) = "Dolares"

    End If

    objWorksheet.Cells(9, 1).Font.bold = True
    objWorksheet.Cells(9, 1).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(9, 1) = "Desde"
    objWorksheet.Cells(9, 2).Select
    Selection.ColumnWidth = 9
    objWorksheet.Cells(9, 2).Font.Size = 4 'aqui tamaño letra
    objWorksheet.Cells(9, 2) = fechai
 
    objWorksheet.Cells(9, 4).Font.bold = True
    objWorksheet.Cells(9, 4).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(9, 4) = "Al"
    objWorksheet.Cells(9, 5).Select
    Selection.ColumnWidth = 9
    objWorksheet.Cells(9, 5).Font.Size = 4 'aqui tamaño letra
    objWorksheet.Cells(9, 5) = fechaf
 
    'fecha a imprimir
    objWorksheet.Cells(9, 6).Font.bold = True
    objWorksheet.Cells(9, 6).Font.Size = 4 'aqui tamaño letra
    objWorksheet.Cells(9, 6) = "FECHA HOY :"
    objWorksheet.Cells(9, 7) = Format(Now, "dd/mm/yyyy") & "- HORA " + Format(Now, "HH:MM:SS")
    objWorksheet.Cells(9, 7).Font.Size = 4 'aqui tamaño letra
 
    'Aqui los margenes cuadrados
    objWorksheet.Range(Cells(10, 1), Cells(10, 12)).Borders.LineStyle = xlContinuous 'Codigo
    objWorksheet.Range(Cells(10, 1), Cells(10, 12)).Interior.color = RGB(215, 215, 215) 'Codigo

    objWorksheet.Cells(10, 1).Font.bold = True
    objWorksheet.Cells(10, 1).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 1) = "Familia"
    objWorksheet.Range("A10").ColumnWidth = 8
  
    objWorksheet.Cells(10, 2).Font.bold = True
    objWorksheet.Cells(10, 2).Font.Size = 7
    objWorksheet.Cells(10, 2) = "Subfamilia"
    objWorksheet.Range("B10").ColumnWidth = 8
  
    objWorksheet.Cells(10, 3).Font.bold = True
    objWorksheet.Cells(10, 3).Font.Size = 7
    objWorksheet.Cells(10, 3) = "Categoria"
    objWorksheet.Range("C10").ColumnWidth = 8
  
    objWorksheet.Cells(10, 4).Font.bold = True
    objWorksheet.Cells(10, 4).Font.Size = 7
    objWorksheet.Cells(10, 4) = "Producto"
    objWorksheet.Range("D10").ColumnWidth = 6
  
    objWorksheet.Cells(10, 5).Font.bold = True
    objWorksheet.Cells(10, 5).Font.Size = 7
    objWorksheet.Cells(10, 5) = "Descripcion"
    objWorksheet.Range("E10").ColumnWidth = 25
  
    objWorksheet.Cells(10, 6).Font.bold = True
    objWorksheet.Cells(10, 6).Font.Size = 7
    objWorksheet.Cells(10, 6) = "Unidad"
    objWorksheet.Range("F10").ColumnWidth = 4
  
    objWorksheet.Cells(10, 7).Font.bold = True
    objWorksheet.Cells(10, 7).Font.Size = 7
    objWorksheet.Cells(10, 7) = "factor"
    objWorksheet.Range("G10").ColumnWidth = 8
  
    objWorksheet.Cells(10, 8).Font.bold = True
    objWorksheet.Cells(10, 8).Font.Size = 7
    objWorksheet.Cells(10, 8) = "Cantidad"
    objWorksheet.Range("H10").ColumnWidth = 5
  
    objWorksheet.Cells(10, 9).Font.bold = True
    objWorksheet.Cells(10, 9).Font.Size = 7
    objWorksheet.Cells(10, 9) = "Saldo"
    objWorksheet.Range("I10").ColumnWidth = 5
  
    objWorksheet.Cells(10, 10).Font.bold = True
    objWorksheet.Cells(10, 10).Font.Size = 7
    objWorksheet.Cells(10, 10) = "Costo"
    objWorksheet.Range("J10").ColumnWidth = 5
  
    objWorksheet.Cells(10, 11).Font.bold = True
    objWorksheet.Cells(10, 11).Font.Size = 7
    objWorksheet.Cells(10, 11) = "Total"
    objWorksheet.Range("K10").ColumnWidth = 5
    
    objWorksheet.Cells(10, 12).Font.bold = True
    objWorksheet.Cells(10, 12).Font.Size = 7
    objWorksheet.Cells(10, 12) = "Minimo"
    objWorksheet.Range("L10").ColumnWidth = 8
  
    Exit Sub
titulo_r_saldo_actual:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

End Sub

Public Sub detalle_reporte_saldo_actual(objWorkBook As Excel.Workbook, _
                                        v As Integer, _
                                        h As Integer, _
                                        local1 As String, _
                                        bodega As String, _
                                        xprecio As Double)
           
    Dim mysql    As String

    Dim mytablez As New ADODB.Recordset

    Dim my_conta As Integer

    Dim my_saldo As Double

    Dim my_total As Double

    Set objWorksheet = objWorkBook.Worksheets(1)

    v = 13
    h = 0

    my_saldo = 0

    mysql = "SELECT pr.familia as familia,pr.producto as producto," & Chr$(10)
    mysql = mysql & "pr.subfamilia as subfamilia, pr.categoria as categoria," & Chr$(10)
    mysql = mysql & "al.unidad as unidad, pr.factor,al.saldo," & Chr$(10)
    mysql = mysql & "al.saldoinicial as saldoinicial, pr.categoria as categoria" & Chr$(10)
    mysql = mysql & "from  almacen al, " & Chr$(10)
    mysql = mysql & "producto pr" & Chr$(10)
    mysql = mysql & "where al.local='" & extra_loquesea(local1) & "' " & Chr$(10)
    mysql = mysql & "and pr.producto = al.PRODUCTO" & Chr$(10)
    mysql = mysql & "order by pr.DESCRIPCIO desc" & Chr$(10)
 
    mytablez.Open mysql, cn, adOpenStatic, adLockOptimistic
 
    If mytablez.RecordCount > 0 Then  'si existe
        Do

            If mytablez.EOF Then Exit Do
            my_conta = mytablez.RecordCount
      
            If temp <> mytablez.Fields("familia") Then
                objWorksheet.Cells(v, h + 1) = mytablez.Fields("familia")
                v = v + 1
                temp = "" & mytablez.Fields("familia")

            End If
     
            If "" & mytablez.Fields("familia") <> temp Then
                temp = "" & mytablez.Fields("familia")
                objWorksheet.Cells(v, h + 2) = "" & mytablez.Fields("familia")
                v = v + 1
            Else
                objWorksheet.Cells(v, h + 2) = "" & mytablez.Fields("producto")
                objWorksheet.Cells(v, h + 3) = "" & mytablez.Fields("subfamilia")
                objWorksheet.Cells(v, h + 4) = "" & mytablez.Fields("unidad")
                objWorksheet.Cells(v, h + 5) = "" & mytablez.Fields("factor")
                objWorksheet.Cells(v, h + 6) = "" & mytablez.Fields("saldo")
      
                my_saldo = my_saldo + mytablez.Fields("saldo")
                objWorksheet.Cells(v, h + 7) = "" & mytablez.Fields("saldoinicial") * xprecio 'esto es el total
        
                my_total = my_total + Val("" & mytablez.Fields("saldoinicial")) * xprecio
                objWorksheet.Cells(v, h + 8) = "" & mytablez.Fields("categoria") 'par aver

                temp1 = "" & mytablez.Fields("familia") & "" & mytablez.Fields("subfamilia")
                sw1 = 1
                v = v + 1
                mytablez.MoveNext

                'inicio 19/0872017 pll
                '        repinv.PB_excel.Value = ((my_conta / v) * 100)
                '
                '        repinv.lbl_excel.Caption = "Actualizando al.." & ((my_conta / v) * 100) & "%"
            End If

        Loop

    End If

    mytablez.Close
    'el total
    objWorksheet.Cells(v, h + 4).Font.bold = True
    objWorksheet.Cells(v, h + 4).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 4) = "TOTAL"
   
    objWorksheet.Cells(v, h + 5).Interior.color = RGB(215, 215, 0)  'resaltador
    objExcel.ActiveSheet.Cells(v, h + 5) = "" & Format(my_saldo, "0.00")
   
    objWorksheet.Cells(v, h + 7).Interior.color = RGB(215, 215, 0)  'resaltador
    objExcel.ActiveSheet.Cells(v, h + 7) = "" & Format(my_total, "0.00")

End Sub

Function busca_familia(buf As String) As String

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "Select * from familia where familia='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_familia = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close
 
End Function

Public Sub titulo_frmChart(fechai As String, fechaf As String)

    On Error GoTo titulo_frmChart

    Set objWorksheet = objWorkBook.Worksheets(1)

    objWorksheet.Cells(9, 1).Font.bold = True
    objWorksheet.Cells(9, 1).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(9, 1) = "Desde"
    objWorksheet.Cells(9, 2).Select
    Selection.ColumnWidth = 9
    objWorksheet.Cells(9, 2).Font.Size = 4 'aqui tamaño letra
    objWorksheet.Cells(9, 2) = fechai
 
    objWorksheet.Cells(9, 4).Font.bold = True
    objWorksheet.Cells(9, 4).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(9, 4) = "Al"
    objWorksheet.Cells(9, 5).Select
    Selection.ColumnWidth = 9
    objWorksheet.Cells(9, 5).Font.Size = 4 'aqui tamaño letra
    objWorksheet.Cells(9, 5) = fechaf
 
    'Aqui los margenes cuadrados
    objWorksheet.Range(Cells(10, 1), Cells(10, 5)).Borders.LineStyle = xlContinuous 'Codigo
    objWorksheet.Range(Cells(10, 1), Cells(10, 5)).Interior.color = RGB(215, 215, 215) 'Codigo

    'fecha a imprimir
    objWorksheet.Cells(9, 6).Font.bold = True
    objWorksheet.Cells(9, 6) = "FECHA HOY :"
    objWorksheet.Cells(9, 7) = Format(Now, "dd/mm/yyyy") & "- HORA " + Format(Now, "HH:MM:SS")
 
    objWorksheet.Cells(10, 1).Font.bold = True
    objWorksheet.Cells(10, 1).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 1) = "Fecha"
    objWorksheet.Range("A10").ColumnWidth = 10
  
    objWorksheet.Cells(10, 2).Font.bold = True
    objWorksheet.Cells(10, 2).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 2) = "Caja Venta"
    objWorksheet.Range("B10").ColumnWidth = 8
  
    objWorksheet.Cells(10, 3).Font.bold = True
    objWorksheet.Cells(10, 3).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 3) = "Total Venta"
    objWorksheet.Range("C10").ColumnWidth = 10
  
    objWorksheet.Cells(10, 4).Font.bold = True
    objWorksheet.Cells(10, 4).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 4) = "Caja Compra"
    objWorksheet.Range("D10").ColumnWidth = 10
  
    objWorksheet.Cells(10, 5).Font.bold = True
    objWorksheet.Cells(10, 5).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 5) = "Total Compras"
    objWorksheet.Range("E10").ColumnWidth = 10
    Exit Sub
titulo_frmChart:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

End Sub

Public Sub titulo_cierre_caja_ciega(objWorkBook As Excel.Workbook)

    Set objWorksheet = objWorkBook.Worksheets(1)

    'fecha a imprimir
    objWorksheet.Cells(11, 8).Font.bold = True
    objWorksheet.Cells(11, 8) = "FECHA HOY :"
    objWorksheet.Cells(11, 9) = Format(Now, "dd/mm/yyyy") & "- HORA " + Format(Now, "HH:MM:SS")
 
    'Aqui los margenes cuadrados
    objWorksheet.Range(Cells(12, 1), Cells(12, 4)).Borders.LineStyle = xlContinuous 'Codigo
    objWorksheet.Range(Cells(12, 1), Cells(12, 4)).Interior.color = RGB(215, 215, 215) 'Codigo

    objWorksheet.Cells(12, 3).Font.bold = True
    objWorksheet.Cells(12, 3) = "Forma Pago"
    objWorksheet.columns("A").ColumnWidth = 20
  
    objWorksheet.Cells(12, 4).Font.bold = True
    objWorksheet.Cells(12, 4) = "Moneda"
    objWorksheet.columns("B").ColumnWidth = 10
  
    objWorksheet.Cells(12, 5).Font.bold = True
    objWorksheet.Cells(12, 5) = "Entrega"
    objWorksheet.columns("B").ColumnWidth = 15
  
    objWorksheet.Cells(12, 6).Font.bold = True
    objWorksheet.Cells(12, 6) = "Caja"
    objWorksheet.columns("C").ColumnWidth = 15
  
End Sub

Public Sub detalle_caja_ciega(objWorkBook As Excel.Workbook, _
                              fecha As String, _
                              caja As String, _
                              turno As String, _
                              v As Integer, _
                              h As Integer)
           
    Dim mysql    As String

    Dim mytablez As New ADODB.Recordset

    Dim my_conta As Integer

    Dim my_saldo As Double

    Dim my_total As Double

    Set objWorksheet = objWorkBook.Worksheets(1)

    v = 13
    h = 0

    my_saldo = 0

    mysql = "SELECT *" & Chr$(10)
    mysql = mysql & "from  cajaciega " & Chr$(10)
    mysql = mysql & "where fecha='" & fecha & "' " & Chr$(10)
    mysql = mysql & "and caja='" & caja & "'" & Chr$(10)
    mysql = mysql & "and turno='" & turno & "'" & Chr$(10)

    mytablez.Open mysql, cn, adOpenStatic, adLockOptimistic
 
    If mytablez.RecordCount > 0 Then  'si existe
        Do

            If mytablez.EOF Then Exit Do
            my_conta = mytablez.RecordCount
            objWorksheet.Cells(v, h + 3) = "" & mytablez.Fields("descripcio")

            If "" & mytablez.Fields("moneda") = "S" Then
                objWorksheet.Cells(v, h + 4) = "Soles"
            Else
                objWorksheet.Cells(v, h + 4) = "Dolares"

            End If

            objWorksheet.Cells(v, h + 5) = "" & mytablez.Fields("entrega")
            objWorksheet.Cells(v, h + 6) = "" & mytablez.Fields("caja")

            If "" & mytablez.Fields("moneda") = "S" Then
                sdx = sdx + Val("" & mytablez.Fields("entrega"))
                sdx2 = sdx2 + Val("" & mytablez.Fields("encaja"))

            End If

            If "" & mytablez.Fields("moneda") = "D" Then
                sdx1 = sdx1 + Val("" & mytablez.Fields("entrega"))
                sdx3 = sdx3 + Val("" & mytablez.Fields("encaja"))

            End If
              
            v = v + 1
            mytablez.MoveNext
            '        repinv.PB_excel.Value = ((my_conta / v) * 100)
            '        repinv.lbl_excel.Caption = "Actualizando al.." & ((my_conta / v) * 100) & "%"
            'End If
        Loop

    End If

    mytablez.Close
    'el total

    objWorksheet.Cells(v, h + 2).Font.bold = True
    objWorksheet.Cells(v, h + 2).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 2) = "TOTAL SOLES"
   
    objWorksheet.Cells(v, h + 3).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 3) = "" & Format(sdx, "0.00")
   
    objWorksheet.Cells(v, h + 4).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 4) = "" & Format(sdx2, "0.00")
      
    v = v + 1

    objWorksheet.Cells(v, h + 2).Font.bold = True
    objWorksheet.Cells(v, h + 2).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 2) = "TOTAL DOLARES"
   
    objWorksheet.Cells(v, h + 3).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 3) = "" & Format(sdx1, "0.00")
   
    objWorksheet.Cells(v, h + 4).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 4) = "" & Format(sdx3, "0.00")
 
End Sub

Public Sub titulo_excel_impresion_total(my_moneda As String)

    On Error GoTo titulo_excel_impresion_total

    Set objWorksheet = objWorkBook.Worksheets(1)

    'tipo de moneda
    objWorksheet.Cells(10, 5) = "Moneda:"
    objWorksheet.Cells(10, 6) = my_moneda
 
    'fecha a imprimir
 
    objWorksheet.Cells(11, 5).Font.bold = True
    objWorksheet.Cells(11, 5) = "FECHA HOY :"
    objWorksheet.Cells(11, 6) = Format(Now, "dd/mm/yyyy") & "- HORA " + Format(Now, "HH:MM:SS")
 
    'Aqui los margenes cuadrados
    objWorksheet.Range(Cells(12, 1), Cells(12, 7)).Borders.LineStyle = xlContinuous 'Codigo
    objWorksheet.Range(Cells(12, 1), Cells(12, 7)).Interior.color = RGB(215, 215, 215) 'Codigo

    objWorksheet.Cells(12, 1).Font.bold = True
    objWorksheet.Cells(12, 1) = "Producto"
    objWorksheet.Range("A12").ColumnWidth = 8
  
    objWorksheet.Cells(12, 2).Font.bold = True
    objWorksheet.Cells(12, 2) = "Descripcion"
    objWorksheet.Range("B12").ColumnWidth = 30
  
    objWorksheet.Cells(12, 3).Font.bold = True
    objWorksheet.Cells(12, 3) = "Unidad"
    objWorksheet.Range("C12").ColumnWidth = 5
  
    objWorksheet.Cells(12, 4).Font.bold = True
    objWorksheet.Cells(12, 4) = "Factor"
    objWorksheet.Range("D12").ColumnWidth = 3
  
    objWorksheet.Cells(12, 5).Font.bold = True
    objWorksheet.Cells(12, 5) = "Cantidad"
    objWorksheet.Range("E12").ColumnWidth = 10
  
    objWorksheet.Cells(12, 6).Font.bold = True
    objWorksheet.Cells(12, 6) = "Precio"
    objWorksheet.Range("F12").ColumnWidth = 12
  
    objWorksheet.Cells(12, 7).Font.bold = True
    objWorksheet.Cells(12, 7) = "Total"
    objWorksheet.Range("G12").ColumnWidth = 12
    Exit Sub
titulo_excel_impresion_total:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

End Sub

Public Sub detalle_excel_impresion_total(my_local As String, _
                                         my_moneda As String, _
                                         fechai As String, _
                                         fechaf As String, _
                                         my_struc_cotizacion_total_excel() As struc_cotizacion_total_excel, _
                                         k As Integer, _
                                         salida As Boolean)
           
    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    Dim my_conta As Integer

    Dim my_saldo As Double

    Dim my_total As Double

    ReDim my_struc_cotizacion_total_excel(0)

    mysql = "SELECT c.PRODUCTO as producto,c.DESCRIPCIO as descripcio," & Chr$(10)
    mysql = mysql & "c.unidad as unidad, c.FACTOR as factor," & Chr$(10)
    mysql = mysql & "c.cantidad as cantidad,c.precio as precio,c.total as total," & Chr$(10)
    mysql = mysql & "p.factor1 as factor1,c.precio as pventa1" & Chr$(10)
    mysql = mysql & "from " & dgusuariog & " c," & Chr$(10)
    mysql = mysql & "precios p" & Chr$(10)
    mysql = mysql & "where c.moneda ='" & my_moneda & "'" & Chr$(10)

    If my_local = "%" Then
        mysql = mysql & "and c.local like'" & my_local & "'" & Chr$(10)
    Else
        mysql = mysql & "and c.local='" & my_local & "'" & Chr$(10)

    End If

    mysql = mysql & "and c.fecha>='" & Format(fechai, "YYYYMMDD") & "'" & Chr$(10)
    mysql = mysql & "and c.fecha<='" & Format(fechaf, "YYYYMMDD") & "' " & Chr$(10)
    mysql = mysql & "and p.PRODUCTO = c.producto" & Chr$(10)
 
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
                ReDim Preserve my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel) + 1)

            End If
     
            If mytablex.Fields("producto") <> "" Then
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).producto = mytablex.Fields("producto")
            Else
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).producto = ""

            End If
      
            If mytablex.Fields("descripcio") <> "" Then
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).descripcion = mytablex.Fields("descripcio")
            Else
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).descripcion = ""

            End If
    
            If mytablex.Fields("unidad") <> "" Then
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).unidad = mytablex.Fields("unidad")
            Else
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).unidad = ""

            End If
    
            If mytablex.Fields("factor") <> "" Then
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).factor = mytablex.Fields("factor")
            Else
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).factor = ""

            End If
   
            If mytablex.Fields("cantidad") <> "" Then
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).cantidad = mytablex.Fields("cantidad")
            Else
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).cantidad = 0

            End If
   
            If mytablex.Fields("precio") <> "" Then
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).precio = mytablex.Fields("precio")
            Else
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).precio = 0

            End If
   
            If mytablex.Fields("total") <> "" Then
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).total = mytablex.Fields("total")
            Else
                my_struc_cotizacion_total_excel(UBound(my_struc_cotizacion_total_excel)).total = 0

            End If
   
            k = k + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

    Exit Sub

End Sub

Public Function titulo_cotizaciones(my_tipo As String, _
                                    my_serie As String, _
                                    my_numero As String, _
                                    my_nombre As String, _
                                    my_codigo As String, _
                                    my_moneda As String, _
                                    my_descrTipo As String, _
                                    my_fecha As String)
                 
    On Error GoTo titulo_cotizaciones

    Set objWorksheet = objWorkBook.Worksheets(1)

    'aqui los marcos
    objWorksheet.Range(Cells(9, 1), Cells(9, 1)).Borders.LineStyle = xlContinuous 'Tipo
    objWorksheet.Range(Cells(9, 1), Cells(9, 1)).Interior.color = RGB(215, 215, 215) 'Tipo
    objWorksheet.Range(Cells(9, 3), Cells(9, 3)).Borders.LineStyle = xlContinuous 'serie
    objWorksheet.Range(Cells(9, 3), Cells(9, 3)).Interior.color = RGB(215, 215, 215) 'serie
    objWorksheet.Range(Cells(9, 5), Cells(9, 5)).Borders.LineStyle = xlContinuous 'numero
    objWorksheet.Range(Cells(9, 5), Cells(9, 5)).Interior.color = RGB(215, 215, 215) 'numero
 
    objWorksheet.Range(Cells(10, 1), Cells(10, 1)).Borders.LineStyle = xlContinuous 'Cliente
    objWorksheet.Range(Cells(10, 1), Cells(10, 1)).Interior.color = RGB(215, 215, 215) 'Cliente
    objWorksheet.Range(Cells(10, 3), Cells(10, 3)).Borders.LineStyle = xlContinuous 'fecha e
    objWorksheet.Range(Cells(10, 3), Cells(10, 3)).Interior.color = RGB(215, 215, 215) 'Nombre
 
    objWorksheet.Range(Cells(10, 5), Cells(10, 5)).Borders.LineStyle = xlContinuous 'Moneda
    objWorksheet.Range(Cells(10, 5), Cells(10, 5)).Interior.color = RGB(215, 215, 215) 'Moneda
 
    objWorksheet.Range(Cells(11, 1), Cells(11, 1)).Borders.LineStyle = xlContinuous 'nombre
    objWorksheet.Range(Cells(11, 1), Cells(11, 1)).Interior.color = RGB(215, 215, 215) 'color interno
 
    'aqui datos del vendedor y el tipo
    objWorksheet.Cells(9, 1) = "Tipo :"
    objWorksheet.Cells(9, 2).Select
    Selection.ColumnWidth = 14
             
    objWorksheet.Cells(9, 2) = my_descrTipo
    objWorksheet.Cells(9, 3) = "Serie:"
    objWorksheet.Range("D9").ColumnWidth = 12
    objWorksheet.Cells(9, 4) = my_serie
 
    objWorksheet.Cells(9, 5) = "Numero:"
    objWorksheet.Cells(9, 6) = my_numero
    objWorksheet.Cells(10, 1) = "Cliente:"
    objWorksheet.Cells(10, 2) = "'" & my_codigo
 
    objWorksheet.Cells(10, 3) = "Fecha E:"
    objWorksheet.Cells(10, 4) = my_fecha
 
    objWorksheet.Cells(10, 5) = "Moneda:"
    objWorksheet.Cells(10, 6) = my_moneda
 
    objWorksheet.Cells(11, 1) = "Nombre:"
    objWorksheet.Cells(11, 2) = my_nombre
 
    objWorksheet.Cells(11, 3).Font.bold = True
    objWorksheet.Cells(11, 3) = "FECHA HOY :"
    objWorksheet.Cells(11, 4) = Format(Now, "dd/mm/yyyy") & "- HORA " + Format(Now, "HH:MM:SS")
 
    '16/08/2017 pll para el trasportista

    objWorksheet.Cells(12, 1).Font.bold = True
    objWorksheet.Cells(12, 1) = "TRANSPORTE"
 
    objWorksheet.Cells(13, 1).Font.bold = True
    objWorksheet.Cells(13, 1) = "Enviar Por:"
 
    objWorksheet.Cells(13, 5).Font.bold = True
    objWorksheet.Cells(13, 5) = "RUC:"
 
    objWorksheet.Cells(14, 1).Font.bold = True
    objWorksheet.Cells(14, 1) = "Lug.Entrega :"
 
    objWorksheet.Cells(14, 5).Font.bold = True
    objWorksheet.Cells(14, 5) = "Fecha Entrega :"
 
    objWorksheet.Cells(15, 1).Font.bold = True
    objWorksheet.Cells(15, 1) = "Destino Final :"
 
    Exit Function
titulo_cotizaciones:

    Select Case Err.Number

        Case 1004
            objWorkBook.Close
            exl.Quit

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

End Function

Public Function detalle_solo_documentos(my_local As String, _
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
                                        my_moneda As String, _
                                        my_struc_solo_documentos() As struc_solo_documentos, _
                                        k As Integer, _
                                        salida As Boolean)
               
    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset
 
    If Len(my_fechai) <> 10 Then Exit Function
    If Len(my_fechaf) <> 10 Then Exit Function
    If Not IsDate(my_fechai) Then Exit Function
    If Not IsDate(my_fechaf) Then Exit Function

    ReDim my_struc_solo_documentos(0)

    mysql = "select * from " & cgusuario & " where " & Chr$(10)

    'If ve = "V" Then
    If acu = "V" Then
        mysql = mysql & "moneda='" & my_moneda & "'" & Chr$(10)
        mysql = mysql & " and fechae>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & " and fechae<='" & Format(my_fechaf, "YYYYMMDD") & "' " & Chr$(10)
    Else
        mysql = mysql & "moneda='" & my_moneda & "'" & Chr$(10) 'aqui
        mysql = mysql & " and fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & " and fecha<='" & Format(my_fechaf, "YYYYMMDD") & "' " & Chr$(10)

    End If

    If Trim(my_local) <> "%" Then
        mysql = mysql & " and local like '" & extra_loquesea(my_local) & "'" & Chr$(10)

    End If

    If Trim(my_tipo) <> "%" Then
        mysql = mysql & " and tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)

    End If

    If Trim(my_caja) <> "%" Then
        mysql = mysql & " and caja like '" & extra_loquesea(my_caja) & "'" & Chr$(10)

    End If

    If Trim(my_turno) <> "%" Then
        mysql = mysql & " and turno like '" & my_turno & "'" & Chr$(10)

    End If

    If my_serie <> "%" Then
        mysql = mysql & " and serie like '" & my_serie & "'" & Chr$(10)

    End If

    If my_numero <> "%" Then
        mysql = mysql & " and numero like '" & my_numero & "'" & Chr$(10)

    End If

    'If codigo <> "%" Then
    '   mysql = mysql & " and codigo like '" & codigo & "'"
    'End If
    'If nombre <> "%" Then
    '   mysql = mysql & " and nombre like '" & nombre & "'"
    'End If
    'If moneda <> "%" Then
    '   mysql = mysql & " and moneda like '" & moneda & "'"
    'End If
    'If estado <> "%" Then
    '   mysql = mysql & " and estado like '" & estado & "'"
    'End If
    'If Trim(PLACA) <> "%" Then
    '   mysql = mysql & " and placa='" & Trim(PLACA) & "'"
    'End If

    If Trim(my_vendedor) <> "%" Then
        mysql = mysql & " and vendedor like '" & my_vendedor & "'" & Chr$(10)

    End If

    If Trim(my_cajero) <> "%" Then
        mysql = mysql & " and usuario like '" & my_cajero & "'" & Chr$(10)

    End If

    If Trim(my_bodega) <> "%" Then
        mysql = mysql & " and bodega like '" & my_bodega & "'" & Chr$(10)

    End If

    If Trim(my_bodegaf) <> "%" Then
        mysql = mysql & " and bodegaf like '" & my_bodegaf & "'" & Chr$(10)

    End If

    If Trim(my_servicio) <> "%" Then
        mysql = mysql & " and  servicio='" & my_servicio & "'" & Chr$(10)

    End If

    'If saldoini.Value = 1 Then
    '  mysql = mysql & " and nop='S' "
    'End If

    If acu <> "C" And acu <> "V" Then
        mysql = mysql & " and acu='" & acu & "'" & Chr$(10)

    End If

    If my_combo2 <> "%" Then
        If my_combo2 = "Atendido" Then
            mysql = mysql & " and  yausado='1'" & Chr$(10)

        End If

        If my_combo2 = "Pendiente" Then
            mysql = mysql & " and  yausado='0'" & Chr$(10)

        End If

    End If

    If acu = "V" Then
        mysql = mysql & " and (acu='A' OR acu='B' OR acu='C' OR acu='D' OR acu='G' )" & Chr$(10)

        If explorap.Check1.Value = 1 Then
            mysql = mysql & " and tipo<>'5'" & Chr$(10)

        End If

    End If

    If acu = "C" Then
        mysql = mysql & " and (acu='J' OR acu='K' OR acu='L' OR acu='M' OR acu='P' )" & Chr$(10)

    End If

    If importacion = "IMPORTACION" Then
        mysql = mysql & " and tipoimp='I'" & Chr$(10)

    End If

    If importacion = "GASTOS" Then
        mysql = mysql & " and tipoimp='G'" & Chr$(10)

    End If

    If importacion = "COMERCIAL" Then
        mysql = mysql & " and (tipoimp='C' or tipoimp is null) " & Chr$(10)

    End If

    If my_ordenado <> "%" Then
        mysql = mysql & "order by " & ordenado & Chr$(10)

    End If

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
                ReDim Preserve my_struc_solo_documentos(UBound(my_struc_solo_documentos) + 1)

            End If
     
            If mytablex.Fields("codigo") <> "" Then
                my_struc_solo_documentos(UBound(my_struc_solo_documentos)).codigo = mytablex.Fields("codigo")
            Else
                my_struc_solo_documentos(UBound(my_struc_solo_documentos)).codigo = ""

            End If
      
            If mytablex.Fields("nombre") <> "" Then
                my_struc_solo_documentos(UBound(my_struc_solo_documentos)).nombre = mytablex.Fields("nombre")
            Else
                my_struc_solo_documentos(UBound(my_struc_solo_documentos)).nombre = ""

            End If
    
            If mytablex.Fields("Local") <> "" Then
                my_struc_solo_documentos(UBound(my_struc_solo_documentos)).local = mytablex.Fields("Local")
            Else
                my_struc_solo_documentos(UBound(my_struc_solo_documentos)).local = ""

            End If
    
            If mytablex.Fields("estado") <> "" Then
                If mytablex.Fields("estado") = "2" Then
                    my_struc_solo_documentos(UBound(my_struc_solo_documentos)).estado = "Cerrado"
                ElseIf mytablex.Fields("estado") = "0" Then
                    my_struc_solo_documentos(UBound(my_struc_solo_documentos)).estado = "Modifica"
                ElseIf mytablex.Fields("estado") = "1" Then
                    my_struc_solo_documentos(UBound(my_struc_solo_documentos)).estado = "Anulado"

                End If

            End If
   
            If mytablex.Fields("tipo") <> "" Then
                my_struc_solo_documentos(UBound(my_struc_solo_documentos)).tipo = mytablex.Fields("tipo")
            Else
                my_struc_solo_documentos(UBound(my_struc_solo_documentos)).tipo = ""

            End If
   
            If mytablex.Fields("serie") <> "" Then
                my_struc_solo_documentos(UBound(my_struc_solo_documentos)).serie = mytablex.Fields("serie")
            Else
                my_struc_solo_documentos(UBound(my_struc_solo_documentos)).serie = ""

            End If
   
            If mytablex.Fields("numero") <> "" Then
                my_struc_solo_documentos(UBound(my_struc_solo_documentos)).Numero = mytablex.Fields("numero")
            Else
                my_struc_solo_documentos(UBound(my_struc_solo_documentos)).Numero = ""

            End If
   
            If mytablex.Fields("fecha") <> "" Then
                my_struc_solo_documentos(UBound(my_struc_solo_documentos)).fecha = mytablex.Fields("fecha")
            Else
                my_struc_solo_documentos(UBound(my_struc_solo_documentos)).fecha = ""

            End If

            If mytablex.Fields("total") <> "" Then
                my_struc_solo_documentos(UBound(my_struc_solo_documentos)).total = mytablex.Fields("total")
            Else
                my_struc_solo_documentos(UBound(my_struc_solo_documentos)).total = 0

            End If
   
            k = k + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
 
    Exit Function
detalle_solo_documentos:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

End Function

Public Function detalle_Fproductos(my_report As String, _
                                   my_agrupacion As String, _
                                   my_local As String, _
                                   my_tipo As String, _
                                   my_serie As String, _
                                   my_numero As String, _
                                   my_codigo As String, _
                                   my_bodega As String, _
                                   my_usuario As String, _
                                   my_servicio As String, _
                                   my_unidad As String, _
                                   my_caja As String, _
                                   my_turno As String, _
                                   my_producto As String, _
                                   my_familia As String, _
                                   my_subfamilia As String, _
                                   my_marca As String, _
                                   my_descripcio As String, _
                                   my_moneda As String, _
                                   my_vendedor As String, _
                                   my_estado As String, _
                                   acu As String, _
                                   my_fechai As String, _
                                   my_fechaf As String, _
                                   my_horai As String, my_horaf As String, my_struc_producto() As struc_producto, salida As Boolean, k As Integer)
       
    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    Dim my_conta As Integer

    Dim my_saldo As Double

    Dim my_total As Double

    ReDim my_struc_producto(0)
 
    mysql = "select " & "d." & my_agrupacion & ",d.Producto," & Chr$(10)
    mysql = mysql & "d.Descripcio,d.moneda as m," & Chr$(10)
    mysql = mysql & "d.UNIDAD," & Chr$(10)
    mysql = mysql & "d.factor," & Chr$(10)
    mysql = mysql & "p.precio," & Chr$(10)
    mysql = mysql & "sum(d.cantidad*d.factor) as xcanti," & Chr$(10)
    mysql = mysql & "sum(d.total) as xtotal,sum(d.tcosto*d.cantidad*d.factor) as xcostoy," & Chr$(10)
    mysql = mysql & "p.COSTOU as xcosto," & Chr$(10)
    mysql = mysql & "(sum(d.total)-sum(d.tcosto*d.cantidad*d.factor)) as xmargen " & Chr$(10)
    mysql = mysql & "from " & my_report & " d," & Chr$(10)
    mysql = mysql & "producto p" & Chr$(10)
    mysql = mysql & "where d.moneda='" & my_moneda & "'" & Chr$(10)
    mysql = mysql & "and d.fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
    mysql = mysql & "and d.fecha<='" & Format(my_fechaf, "YYYYMMDD") & "' " & Chr$(10)

    If my_local <> "%" Then
        mysql = mysql & "and d.local='" & extra_loquesea(my_local) & "'" & Chr$(10)

    End If

    If my_tipo <> "%" Then
        mysql = mysql & "and d.tipo like '" & my_tipo & "'" & Chr$(10)

    End If

    If my_serie <> "%" Then
        mysql = mysql & "and d.serie like '" & my_serie & "'" & Chr$(10)

    End If

    If my_numero <> "%" Then
        mysql = mysql & " and d.numero like '" & my_numero & "'" & Chr$(10)

    End If

    If my_codigo <> "%" Then
        mysql = mysql & " and d.codigo like '" & my_codigo & "'" & Chr$(10)

    End If

    If my_usuario <> "%" Then
        mysql = mysql & " and d.usuario like '" & my_usuario & "'" & Chr$(10)

    End If

    If my_bodega <> "%" Then
        mysql = mysql & " and d.bodega like '" & extra_loquesea(my_bodega) & "'" & Chr$(10)

    End If

    If my_servicio <> "%" Then
        mysql = mysql & " and d.servicio='" & extra_loquesea(my_servicio) & "'" & Chr$(10)

    End If

    If my_unidad <> "%" Then
        mysql = mysql & " and d.unidad like '" & my_unidad & "'" & Chr$(10)

    End If

    If my_caja <> "%" Then
        mysql = mysql & " and d.caja like '" & my_caja & "'" & Chr$(10)

    End If

    If my_turno <> "%" Then
        mysql = mysql & " and d.turno like '" & my_turno & "'" & Chr$(10)

    End If

    If my_producto <> "%" Then
        mysql = mysql & " and d.producto like '" & my_producto & "'" & Chr$(10)

    End If

    If my_horai <> "%" And horaf <> "%" Then
        mysql = mysql & " and d.HORA BETWEEN '" & my_horai & "' AND '" & my_horaf & "'" & Chr$(10)

    End If

    If my_familia <> "%" Then
        mysql = mysql & " and d.familia like '" & my_familia & "'" & Chr$(10)

    End If

    If my_subfamilia <> "%" Then
        mysql = mysql & " and d.subfamilia like '" & my_subfamilia & "'" & Chr$(10)

    End If

    If my_marca <> "%" Then
        mysql = mysql & " and d.marca like '" & my_marca & "'" & Chr$(10)

    End If

    If my_descripcio <> "%" Then
        mysql = mysql & " and d.descripcio like '" & my_descripcio & "'" & Chr$(10)

    End If

    If my_moneda <> "%" Then
        mysql = mysql & " and d.moneda like '" & my_moneda & "'" & Chr$(10)

    End If

    If my_vendedor <> "%" Then
        mysql = mysql & " and d.vendedor like '" & my_vendedor & "'" & Chr$(10)

    End If

    If my_estado <> "%" Then
        mysql = mysql & " and d.estado='" & my_estado & "'" & Chr$(10)

    End If

    If acu = "V" Then
        mysql = mysql & " and (d.acu='1' or d.acu='A' or d.acu='B' or d.acu='C' or d.acu='D' or d.acu='G' or d.acu='E' or d.acu='F') " & Chr$(10)

    End If

    If acu = "C" Then
        mysql = mysql & " and (d.acu='J' or d.acu='K' or d.acu='L' or d.acu='M' or d.acu='P' or d.acu='N' or d.acu='O') " & Chr$(10)

    End If

    If my_horai <> "%" And my_horaf <> "%" Then
        mysql = mysql & " and d.HORA BETWEEN '" & my_horai & "' AND '" & my_horaf & "'" & Chr$(10)

    End If

    ybuf = " SUM(d.cantidad*d.factor) "
    mysql = mysql & "and p.producto = d.PRODUCTO" & Chr$(10)
    mysql = mysql & "group by " & "d." & my_agrupacion & ", d.producto,d.Descripcio,d.moneda,p.COSTOU, d.UNIDAD,d.factor,p.precio " & Chr$(10)
    mysql = mysql & "order by " & "d." & my_agrupacion & " ," & ybuf & " DESC "

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
                ReDim Preserve my_struc_producto(UBound(my_struc_producto) + 1)

            End If

            If my_agrupacion = "Familia" Then
                If mytablex.Fields("familia") <> my_familia Then
                    my_struc_producto(UBound(my_struc_producto)).familia = mytablex.Fields("familia")
                    my_familia = mytablex.Fields("familia")
                Else
                    my_struc_producto(UBound(my_struc_producto)).familia = ""
                    my_familia = mytablex.Fields("familia")

                End If

            End If
     
            If mytablex.Fields("producto") <> "" Then
                my_struc_producto(UBound(my_struc_producto)).producto = mytablex.Fields("producto")
            Else
                my_struc_producto(UBound(my_struc_producto)).producto = ""

            End If

            If mytablex.Fields("Descripcio") <> "" Then
                my_struc_producto(UBound(my_struc_producto)).descripcion = mytablex.Fields("Descripcio")
            Else
                my_struc_producto(UBound(my_struc_producto)).descripcion = ""

            End If

            If mytablex.Fields("xcanti") <> "" Then
                my_struc_producto(UBound(my_struc_producto)).xcanti = mytablex.Fields("xcanti")
            Else
                my_struc_producto(UBound(my_struc_producto)).xcanti = 0

            End If

            '**
            If mytablex.Fields("xcosto") <> "" Then
                my_struc_producto(UBound(my_struc_producto)).xcostou = mytablex.Fields("xcosto")
            Else
                my_struc_producto(UBound(my_struc_producto)).xcostou = 0

            End If

            '//
            If mytablex.Fields("xtotal") <> "" Then
                my_struc_producto(UBound(my_struc_producto)).xtotal = mytablex.Fields("xtotal")
            Else
                my_struc_producto(UBound(my_struc_producto)).xtotal = 0

            End If

            If mytablex.Fields("m") <> "" Then
                my_struc_producto(UBound(my_struc_producto)).moneda = mytablex.Fields("m")
            Else
                my_struc_producto(UBound(my_struc_producto)).moneda = ""

            End If

            If mytablex.Fields("unidad") <> "" Then
                my_struc_producto(UBound(my_struc_producto)).unidad = mytablex.Fields("unidad")
            Else
                my_struc_producto(UBound(my_struc_producto)).unidad = ""

            End If

            If mytablex.Fields("factor") <> "" Then
                my_struc_producto(UBound(my_struc_producto)).factor = mytablex.Fields("factor")
            Else
                my_struc_producto(UBound(my_struc_producto)).factor = ""

            End If

            If mytablex.Fields("precio") <> "" Then
                my_struc_producto(UBound(my_struc_producto)).precio = mytablex.Fields("precio")
            Else
                my_struc_producto(UBound(my_struc_producto)).precio = 0

            End If
    
            k = k + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

End Function

Public Sub detalle_forma_pago(my_local As String, _
                              my_tipo As String, _
                              my_serie As String, _
                              my_numero As String, _
                              my_vpfago As String, _
                              my_codigo As String, _
                              my_nombre As String, _
                              my_moneda As String, _
                              my_observa As String, _
                              my_cajero As String, _
                              my_caja As String, _
                              my_turno As String, _
                              my_concepto As String, _
                              my_subconcepto As String, _
                              my_estado As String, _
                              my_agrupamiento, _
                              my_fechai As String, _
                              my_fechaf As String, _
                              my_vendedor As String, _
                              acu As String)
           
    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo detalle_forma_pago

    Set objWorksheet = objWorkBook.Worksheets(1)

    If Len(my_fechai) <> 10 Then Exit Sub
    If Len(my_fechaf) <> 10 Then Exit Sub
    If Not IsDate(my_fechai) Then Exit Sub
    If Not IsDate(my_fechaf) Then Exit Sub

    v = 11
    h = 1

    mysql = "select * from fpagov where " & Chr$(10)
    mysql = mysql & "fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
    mysql = mysql & "and fecha<='" & Format(my_fechaf, "YYYYMMDD") & "' " & Chr$(10)

    If my_tipo <> "%" Then
        mysql = mysql & "and tipo like '" & my_tipo & "'" & Chr$(10)

    End If

    If my_vpfago <> "%" Then
        mysql = mysql & "and fpago like '" & my_vpfago & "'" & Chr$(10)

    End If

    If my_local <> "%" Then
        mysql = mysql & "and local like '" & my_local & "'" & Chr$(10)

    End If

    If my_vendedor <> "%" Then
        mysql = mysql & "and vendedor like '" & my_vendedor & "'" & Chr$(10)

    End If

    ''03/07/2017 KENYO 'Mejora Reporte de Formas de Pago (TC, Propina adicional 14/09/2017
  
    If acu <> "%" Then
        If acu = "V" Then
            mysql = mysql & " and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') " & Chr$(10)

        End If

        If acu = "C" Then
            mysql = mysql & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') " & Chr$(10)

        End If

        If acu = "R" Then
            mysql = mysql & " and (acu='V' or acu='W') " & Chr$(10)

        End If

    End If

    ''03/07/2017 KENYO 'Mejora Reporte de Formas de Pago (TC, Propina adicional  14/09/2017
    If my_numero <> "%" Then
        mysql = mysql & " and numero like '" & my_numero & "'" & Chr$(10)

    End If

    If my_codigo <> "%" Then
        mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)

    End If

    If my_nombre <> "%" Then
        mysql = mysql & "and nombre like '" & my_nombre & "'" & Chr$(10)

    End If

    If my_moneda <> "%" Then
        mysql = mysql & "and moneda like '" & my_moneda & "'" & Chr$(10)

    End If

    If my_observa <> "%" Then
        mysql = mysql & "and observa like '" & my_observa & "'" & Chr$(10)

    End If

    If my_cajero <> "%" Then
        mysql = mysql & "and usuario like '" & my_cajero & "'" & Chr$(10)

    End If

    If my_caja <> "%" Then
        mysql = mysql & "and caja like '" & my_caja & "'" & Chr$(10)

    End If

    If my_turno <> "%" Then
        mysql = mysql & " and turno like '" & my_turno & "'" & Chr$(10)

    End If

    If my_concepto <> "%" Then
        mysql = mysql & "and concepto like '" & my_concepto & "'" & Chr$(10)

    End If

    If my_subconcepto <> "%" Then
        mysql = mysql & "and subconcepto like '" & my_subconcepto & "'" & Chr$(10)

    End If

    If my_estado <> "%" Then
        mysql = mysql & "and estado like '" & my_estado & "'" & Chr$(10)

    End If

    'aqui

    mysql = mysql & " order by  fecha,fpago,numero" & Chr$(10)
    'aqui

    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF Then
        salida = False
        Exit Sub
    Else
        salida = True
        mytablex.MoveFirst

        Do Until mytablex.EOF

            If mytablex.RecordCount = 1 Then
                If mytablex.Fields("MONEDA") = "S" Then
                    objWorksheet.Cells(9, 10).Font.bold = True
                    objWorksheet.Cells(9, 10).Font.Size = 5
                    objWorksheet.Cells(9, 10).Select
                    Selection.ColumnWidth = 9
                  
                    objWorksheet.Cells(9, 10) = "Moneda"
                    objWorksheet.Cells(9, 11).Font.Size = 5
                    objWorksheet.Cells(9, 11) = "Soles"
                    '
                    '         objWorksheet.Cells(10, 11).Font.Size = 4
         
                Else
                    objWorksheet.Cells(9, 9).Font.bold = True
                    objWorksheet.Cells(9, 9) = "Moneda"
                    objWorksheet.Cells(9, 10) = "Dolares"

                End If

            End If

            If mytablex.Fields("TIPO") = "FC" Then
                objWorksheet.Cells(v, h) = "Factura Compra"
            ElseIf mytablex.Fields("TIPO") = "BC" Then
                objWorksheet.Cells(v, h) = "Boleta Compra"
            ElseIf mytablex.Fields("TIPO") = "1" Then
                objWorksheet.Cells(v, h) = "Factura"
            ElseIf mytablex.Fields("TIPO") = "2" Then
                objWorksheet.Cells(v, h) = "Boleta"

            End If

            objWorksheet.Cells(v, h + 1).Font.Size = 7 'aqui tamaño letra pll
            objWorksheet.Cells(v, h + 1) = "'" & mytablex.Fields("SERIE")
            objWorksheet.Cells(v, h + 2).Font.Size = 7 'aqui tamaño letra pll
            objWorksheet.Cells(v, h + 2) = "'" & mytablex.Fields("NUMERO")
            objWorksheet.Cells(v, h + 3).Font.Size = 7 'aqui tamaño letra pll
            objWorksheet.Cells(v, h + 3) = "'" & mytablex.Fields("FECHA")
            objWorksheet.Cells(v, h + 4).Font.Size = 7 'aqui tamaño letra pll
            objWorksheet.Cells(v, h + 4) = "" & Format(mytablex.Fields("TOTAL"), "0.00")
    
            my_total = my_total + Format(Val("" & mytablex.Fields("total")), "0.00")
            objWorksheet.Cells(v, h + 5).Font.Size = 7 'aqui tamaño letra pll
            objWorksheet.Cells(v, h + 5) = "" & mytablex.Fields("FPAGO")
            objWorksheet.Cells(v, h + 6).Font.Size = 7 'aqui tamaño letra pll
            objWorksheet.Cells(v, h + 6) = "" & mytablex.Fields("DESCRIPCIO")
            objWorksheet.Cells(v, h + 7).Font.Size = 7 'aqui tamaño letra pll
            objWorksheet.Cells(v, h + 7) = "" & mytablex.Fields("ORDEN")

            'PARA EL ESTADO DESCRIPCION
            If mytablex.Fields("estado") = 2 Then
                objWorksheet.Cells(v, h + 8).Font.Size = 7 'aqui tamaño letra pll
                objWorksheet.Cells(v, h + 8) = "Cerrado"
            ElseIf mytablex.Fields("estado") = 0 Then
                objWorksheet.Cells(v, h + 8).Font.Size = 7 'aqui tamaño letra pll
                objWorksheet.Cells(v, h + 8) = "Modifica"
            ElseIf mytablex.Fields("estado") = 1 Then
                objWorksheet.Cells(v, h + 8).Font.Size = 7 'aqui tamaño letra pll
                objWorksheet.Cells(v, h + 8) = "Anulado"

            End If

            objWorksheet.Cells(v, h + 9).Font.Size = 7 'aqui tamaño letra pll
            objWorksheet.Cells(v, h + 9) = "" & mytablex.Fields("CODIGO")
            objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra pll
            objWorksheet.Cells(v, h + 10) = "" & mytablex.Fields("USUARIO")
            objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra pll
            objWorksheet.Cells(v, h + 11) = "" & mytablex.Fields("VENDEDOR")
            objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letra pll
            objWorksheet.Cells(v, h + 12) = "" & mytablex.Fields("CAJA")
            objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra pll
            objWorksheet.Cells(v, h + 13) = "" & mytablex.Fields("TURNO")
    
            v = v + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
    'aqui los totales
    objWorksheet.Cells(v, h + 3).Font.bold = True
    objWorksheet.Cells(v, h + 3).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 3) = "TOTAL"
   
    objWorksheet.Cells(v, h + 4).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 4) = "" & Format(my_total, "0.00") 'el total
 
    Exit Sub
detalle_forma_pago:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

End Sub

Public Sub R_Saldo_actual(my_producto As String, _
                          my_barras As String, _
                          my_descripcio As String, _
                          my_familia As String, _
                          my_subfamilia As String, _
                          my_seccion As String, _
                          my_categoria As String, _
                          my_linea As String, _
                          my_color As String, _
                          my_marca As String, _
                          my_igv As String, _
                          my_moneda As String, _
                          my_proveedor As String, _
                          my_local1 As String, _
                          my_bodega As String, _
                          my_monedac As String, _
                          my_fechai As String, _
                          my_fechaf As String, _
                          my_conigv As String, _
                          my_gcanti As String, _
                          my_quecosto As String, _
                          my_fechavi As String, _
                          my_fechavf As String, _
                          my_fechavpi As String, _
                          my_fechavpf As String, my_fechari As String, my_fecharf As String, my_saldo As String, my_struc_saldo_actual() As struc_saldo_actual, salida As Boolean, k As Integer)

    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    ReDim my_struc_saldo_actual(0)
    'inicio 26/09/20167 pll
    'mysql = "select distinct p.familia,p.producto,p.DESCRIPCIO," & Chr$(10)
    'mysql = mysql & "a.saldo as cantidad ,p.FECHAVENCE," & Chr$(10)
    'mysql = mysql & "p.costop,p.costou,p.monedac" & Chr$(10)
    'mysql = mysql & "from producto p, " & Chr$(10)
    'mysql = mysql & "almacen a" & Chr$(10)
    'mysql = mysql & "where MONEDAC ='" & my_moneda & "'" & Chr$(10)
    'mysql = mysql & "and a.PRODUCTO = p.producto" & Chr$(10)
    'mysql = mysql & "and a.local='" & my_local1 & "'" & Chr$(10)
    'mysql = mysql & "and a.saldo > 0" & Chr$(10)
    mysql = ""
    mysql = "select distinct p.familia," & Chr$(10)
    mysql = mysql & "p.subfamilia,p.CATEGORIA,p.minimo," & Chr$(10)
    mysql = mysql & "p.producto , p.descripcio, " & Chr$(10)
    mysql = mysql & "p.unidad as unidad,p.factor as factor," & Chr$(10)
    mysql = mysql & "a.saldo as cantidad ," & Chr$(10)
    mysql = mysql & "p.costop,p.costou,p.monedac" & Chr$(10)
    mysql = mysql & "from producto p," & Chr$(10)
    mysql = mysql & "almacen a" & Chr$(10)
    mysql = mysql & "where p.producto like '" & my_producto & "'" & Chr$(10)
    'mysql = mysql & "and MONEDAC ='" & my_moneda & "'" & Chr$(10)
    mysql = mysql & "and a.local='" & my_local1 & "'" & Chr$(10)
    mysql = mysql & "and a.PRODUCTO = p.producto" & Chr$(10)

    If Mid(my_saldo, 6, 2) <> Null Then
        mysql = mysql & "and a.saldo" & Mid(my_saldo, 6, 2) & Chr$(10)

    End If

    'fin 26/09/20167 pll
    'inicio 26/09/20167 pll
    'If my_producto <> "%" Then
    '  mysql = mysql & "and producto like '" & my_producto & "'" & Chr$(10)
    'End If
    'fin 26/09/20167 pll

    If my_barras <> "%" Then
        mysql = mysql & " and barras like '" & my_barras & "'" & Chr$(10)

    End If

    If my_descripcio <> "%" Then
        mysql = mysql & " and descripcio like '" & my_descripcio & "'" & Chr$(10)

    End If

    If my_familia <> "%" Then
        mysql = mysql & " and familia like '" & extra_loquesea(my_familia) & "'" & Chr$(10)

    End If

    If my_subfamilia <> "%" Then
        mysql = mysql & " and subfamilia like '" & my_subfamilia & "'" & Chr$(10)

    End If

    If my_seccion <> "%" Then
        mysql = mysql & " and seccion like '" & my_seccion & "'" & Chr$(10)

    End If

    If my_categoria <> "%" Then
        mysql = mysql & " and categoria like '" & my_categoria & "'" & Chr$(10)

    End If

    If my_linea <> "%" Then
        mysql = mysql & " and linea like '" & my_linea & "'" & Chr$(10)

    End If

    If my_color <> "%" Then
        mysql = mysql & " and color like '" & my_color & "'" & Chr$(10)

    End If

    If my_marca <> "%" Then
        mysql = mysql & " and marca like '" & my_marca & "'" & Chr$(10)

    End If

    If my_fechavi <> "%" And my_fechavf <> "%" Then
        If IsDate(my_fechavi) And IsDate(my_fechavf) Then
            mysql = mysql & "  fechavence>='" & Format(my_fechavi, "YYYYMMDD") & "'" & Chr$(10)
            mysql = mysql & " and fechavence<='" & Format(my_fechavf, "YYYYMMDD") & "' " & Chr$(10)

        End If

    End If

    If my_igv = "EXENTO" Then
        mysql = mysql & " and igv=0" & Chr$(10)

    End If

    If igv = "GRAVADO" Then
        mysql = mysql & " and igv>0" & Chr$(10)

    End If

    mysql = mysql & " order by p.familia,p.producto desc" & Chr$(10)

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
                ReDim Preserve my_struc_saldo_actual(UBound(my_struc_saldo_actual) + 1)

            End If

            If k = 0 Then
                my_familia = mytablex.Fields("familia")
                my_struc_saldo_actual(UBound(my_struc_saldo_actual)).familia = my_familia
            Else

                If my_familia <> mytablex.Fields("familia") Then
                    my_struc_saldo_actual(UBound(my_struc_saldo_actual)).familia = mytablex.Fields("familia")
                    my_familia = mytablex.Fields("familia")
                Else
                    my_struc_saldo_actual(UBound(my_struc_saldo_actual)).familia = ""

                End If

            End If

            'my_familia = mytablex.Fields("familia")
            'If my_struc_saldo_actual(UBound(my_struc_saldo_actual)).familia <> mytablex.Fields("familia") Then
            '    my_struc_saldo_actual(UBound(my_struc_saldo_actual)).familia = mytablex.Fields("familia")
            'Else
            '  my_struc_saldo_actual(UBound(my_struc_saldo_actual)).familia = ""
            'End If
            If mytablex.Fields("subfamilia") <> "" Then
                my_struc_saldo_actual(UBound(my_struc_saldo_actual)).subfamilia = mytablex.Fields("subfamilia")
            Else
                my_struc_saldo_actual(UBound(my_struc_saldo_actual)).subfamilia = ""

            End If

            If mytablex.Fields("categoria") <> "" Then
                my_struc_saldo_actual(UBound(my_struc_saldo_actual)).categoria = mytablex.Fields("categoria")
            Else
                my_struc_saldo_actual(UBound(my_struc_saldo_actual)).categoria = ""

            End If

            If mytablex.Fields("producto") <> "" Then
                my_struc_saldo_actual(UBound(my_struc_saldo_actual)).producto = mytablex.Fields("producto")
            Else
                my_struc_saldo_actual(UBound(my_struc_saldo_actual)).producto = ""

            End If

            If mytablex.Fields("Descripcio") <> "" Then
                my_struc_saldo_actual(UBound(my_struc_saldo_actual)).descripcion = mytablex.Fields("Descripcio")
            Else
                my_struc_saldo_actual(UBound(my_struc_saldo_actual)).descripcion = ""

            End If

            If mytablex.Fields("unidad") <> "" Then
                my_struc_saldo_actual(UBound(my_struc_saldo_actual)).unidad = mytablex.Fields("unidad")
            Else
                my_struc_saldo_actual(UBound(my_struc_saldo_actual)).unidad = ""

            End If

            If mytablex.Fields("factor") <> "" Then
                my_struc_saldo_actual(UBound(my_struc_saldo_actual)).factor = mytablex.Fields("factor")
            Else
                my_struc_saldo_actual(UBound(my_struc_saldo_actual)).factor = ""

            End If

            If mytablex.Fields("cantidad") <> "" Then
                my_struc_saldo_actual(UBound(my_struc_saldo_actual)).cantidad = mytablex.Fields("cantidad")
                my_cantidad = mytablex.Fields("cantidad")
            Else
                my_struc_saldo_actual(UBound(my_struc_saldo_actual)).cantidad = 0

            End If

            If mytablex.Fields("costou") <> "" Then
                my_struc_saldo_actual(UBound(my_struc_saldo_actual)).costou = mytablex.Fields("costou") 'costo
                my_costou = mytablex.Fields("costou")
            Else
                my_struc_saldo_actual(UBound(my_struc_saldo_actual)).costou = 0

            End If

            ' If mytablex.Fields("cantidad") * mytablex.Fields("costou") <> "" Then
      
            'If mytablex.Fields("cantidad") <> "" Then
            my_totalxx = my_cantidad * my_costou
            '   MsgBox "my_ytotal" & my_totalxx
            my_struc_saldo_actual(UBound(my_struc_saldo_actual)).total = my_totalxx

            'Else
            '  my_struc_saldo_actual(UBound(my_struc_saldo_actual)).total = 0
            'End If
            If mytablex.Fields("minimo") <> "" Then
                my_struc_saldo_actual(UBound(my_struc_saldo_actual)).minimo = mytablex.Fields("minimo")
            Else
                my_struc_saldo_actual(UBound(my_struc_saldo_actual)).minimo = 0

            End If

            If mytablex.Fields("monedac") <> "" Then
                my_struc_saldo_actual(UBound(my_struc_saldo_actual)).moneda = mytablex.Fields("monedac")
            Else
                my_struc_saldo_actual(UBound(my_struc_saldo_actual)).moneda = 0

            End If

            k = k + 1
            ' mytablex.MoveNext
            'Loop
            ' End If
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
 
End Sub

Public Sub Detalle_saldo_actual(my_local1 As String, _
                                my_bodega As String, _
                                my_struc_saldo_actual() As struc_saldo_actual, _
                                k As Integer)

    'aqui tengo que investigar pll
    Dim mysql    As String

    Dim mytabley As New ADODB.Recordset

    v = 11
    h = 0
    sdx = 0
    sdx1 = 0
    sdx2 = 0

    On Error GoTo Detalle_saldo_actual

    Set objWorksheet = objWorkBook.Worksheets(1)

    For j = 0 To k - 1
        mysql = "SELECT pr.FAMILIA as familia,p.UNIDAD1,p.UNIDAD2,p.UNIDAD3," & Chr$(10)
        mysql = mysql & "p.unidad4 , p.unidad5, " & Chr$(10)
        mysql = mysql & "p.UNIDAD6,p.UNIDAD7,p.UNIDAD7,p.UNIDAD8,p.UNIDAD9,p.UNIDAD10," & Chr$(10)
        mysql = mysql & "p.factor1,p.factor2,p.factor3,p.factor4,p.factor5," & Chr$(10)
        mysql = mysql & "p.factor6,p.factor7,p.factor8,p.factor9,p.factor10," & Chr$(10)
        mysql = mysql & "p.pventa1,p.pventa2,p.pventa3,p.pventa4,p.pventa5," & Chr$(10)
        mysql = mysql & "p.pventa6,p.pventa7,p.pventa8,p.pventa9,p.pventa10,d.SUBTOTAL" & Chr$(10)
        mysql = mysql & "from precios p," & Chr$(10)
        mysql = mysql & "producto pr," & Chr$(10)
        mysql = mysql & "detalle d" & Chr$(10)
        mysql = mysql & "where p.producto='" & "" & my_struc_saldo_actual(j).producto & "'" & Chr$(10)
        mysql = mysql & "AND pr.producto = p.producto" & Chr$(10)
        mysql = mysql & "and d.PRODUCTO = p.PRODUCTO" & Chr$(10)
        mysql = mysql & "AND D.FLAGE ='E'" & Chr$(10)
        mysql = mysql & "and p.local= '" & my_local1 & "'" & Chr$(10)
        mysql = mysql & "Group By pr.FAMILIA," & Chr$(10)
        mysql = mysql & "p.UNIDAD1,p.UNIDAD2,p.UNIDAD3," & Chr$(10)
        mysql = mysql & "p.UNIDAD4,p.UNIDAD5,p.UNIDAD6," & Chr$(10)
        mysql = mysql & "p.UNIDAD7,p.UNIDAD8,p.UNIDAD9," & Chr$(10)
        mysql = mysql & "p.UNIDAD10," & Chr$(10)
        mysql = mysql & "p.FACTOR1,p.FACTOR2,p.FACTOR3," & Chr$(10)
        mysql = mysql & "p.FACTOR4,p.FACTOR5,p.FACTOR6," & Chr$(10)
        mysql = mysql & "p.FACTOR7,p.FACTOR8,p.FACTOR9,p.FACTOR10," & Chr$(10)
        mysql = mysql & "p.PVENTA1,p.PVENTA2,p.PVENTA3,p.PVENTA4," & Chr$(10)
        mysql = mysql & "p.PVENTA5,p.PVENTA6,p.PVENTA7,p.PVENTA8," & Chr$(10)
        mysql = mysql & "p.pventa9 , p.pventa10,d.SUBTOTAL" & Chr$(10)
        mytabley.Open mysql, cn, adOpenStatic, adLockOptimistic
 
        If mytabley.RecordCount > 0 Then  'si existe
            If mytabley.Fields("familia") <> my_familia Then
                objWorksheet.Cells(v, h + 1) = mytabley.Fields("familia")
                my_familia = mytabley.Fields("familia")

            End If

            objWorksheet.Cells(v, h + 2) = my_struc_saldo_actual(j).producto
            objWorksheet.Cells(v, h + 3) = "'" & my_struc_saldo_actual(j).descripcion

            If mytabley.Fields("unidad1") <> "" Then
                objWorksheet.Cells(v, h + 4) = "" & mytabley.Fields("unidad1")
            ElseIf mytabley.Fields("unidad2") <> "" Then
                objWorksheet.Cells(v, h + 4) = "" & mytabley.Fields("unidad2")
            ElseIf mytabley.Fields("unidad3") <> "" Then
                objWorksheet.Cells(v, h + 4) = "" & mytabley.Fields("unidad3")
            ElseIf mytabley.Fields("unidad4") <> "" Then
                objWorksheet.Cells(v, h + 4) = "" & mytabley.Fields("unidad4")
            ElseIf mytabley.Fields("unidad5") <> "" Then
                objWorksheet.Cells(v, h + 4) = "" & mytabley.Fields("unidad5")
            ElseIf mytabley.Fields("unidad6") <> "" Then
                objWorksheet.Cells(v, h + 4) = "" & mytabley.Fields("unidad6")
            ElseIf mytabley.Fields("unidad7") <> "" Then
                objWorksheet.Cells(v, h + 4) = "" & mytabley.Fields("unidad7")
            ElseIf mytabley.Fields("unidad8") <> "" Then
                objWorksheet.Cells(v, h + 4) = "" & mytabley.Fields("unidad8")
            ElseIf mytabley.Fields("unidad9") <> "" Then
                objWorksheet.Cells(v, h + 4) = "" & mytabley.Fields("unidad9")
            ElseIf mytabley.Fields("unidad10") <> "" Then
                objWorksheet.Cells(v, h + 4) = "" & mytabley.Fields("unidad10")

            End If
   
            If mytabley.Fields("factor1") <> "" Then
                objWorksheet.Cells(v, h + 5) = "" & mytabley.Fields("factor1")
                xfactor = "" & mytabley.Fields("factor1")
            ElseIf mytabley.Fields("factor2") <> "" Then
                objWorksheet.Cells(v, h + 5) = "" & mytabley.Fields("factor2")
                xfactor = "" & mytabley.Fields("factor2")
            ElseIf mytabley.Fields("factor3") <> "" Then
                objWorksheet.Cells(v, h + 5) = "" & mytabley.Fields("factor3")
                xfactor = "" & mytabley.Fields("factor3")
            ElseIf mytabley.Fields("factor4") <> "" Then
                objWorksheet.Cells(v, h + 5) = "" & mytabley.Fields("factor4")
                xfactor = "" & mytabley.Fields("factor4")
            ElseIf mytabley.Fields("factor5") <> "" Then
                objWorksheet.Cells(v, h + 5) = "" & mytabley.Fields("factor5")
                xfactor = "" & mytabley.Fields("factor5")
            ElseIf mytabley.Fields("factor6") <> "" Then
                objWorksheet.Cells(v, h + 5) = "" & mytabley.Fields("factor6")
                xfactor = "" & mytabley.Fields("factor6")
            ElseIf mytabley.Fields("factor7") <> "" Then
                objWorksheet.Cells(v, h + 5) = "" & mytabley.Fields("factor7")
                xfactor = "" & mytabley.Fields("factor7")
            ElseIf mytabley.Fields("factor8") <> "" Then
                objWorksheet.Cells(v, h + 5) = "" & mytabley.Fields("factor8")
                xfactor = "" & mytabley.Fields("factor8")
            ElseIf mytabley.Fields("factor9") <> "" Then
                objWorksheet.Cells(v, h + 5) = "" & mytabley.Fields("factor9")
                xfactor = "" & mytabley.Fields("factor9")
            ElseIf mytabley.Fields("factor10") <> "" Then
                objWorksheet.Cells(v, h + 5) = "" & mytabley.Fields("factor10")
                xfactor = "" & mytabley.Fields("factor10")

            End If

            If mytabley.Fields("pventa1") <> "" Then
                objWorksheet.Cells(v, h + 6) = "" & mytabley.Fields("pventa1")
                Punitario = "" & mytabley.Fields("pventa1")
            ElseIf mytabley.Fields("pventa2") <> "" Then
                objWorksheet.Cells(v, h + 6) = "" & mytabley.Fields("pventa2")
                Punitario = "" & mytabley.Fields("pventa2")
            ElseIf mytabley.Fields("pventa3") <> "" Then
                objWorksheet.Cells(v, h + 6) = "" & mytabley.Fields("pventa3")
                Punitario = "" & mytabley.Fields("pventa3")
            ElseIf mytabley.Fields("pventa4") <> "" Then
                objWorksheet.Cells(v, h + 6) = "" & mytabley.Fields("pventa4")
                Punitario = "" & mytabley.Fields("pventa4")
            ElseIf mytabley.Fields("pventa5") <> "" Then
                objWorksheet.Cells(v, h + 6) = "" & mytabley.Fields("pventa5")
                Punitario = "" & mytabley.Fields("pventa5")
            ElseIf mytabley.Fields("pventa6") <> "" Then
                objWorksheet.Cells(v, h + 6) = "" & mytabley.Fields("pventa6")
                Punitario = "" & mytabley.Fields("pventa6")
            ElseIf mytabley.Fields("pventa7") <> "" Then
                objWorksheet.Cells(v, h + 6) = "" & mytabley.Fields("pventa7")
                Punitario = "" & mytabley.Fields("pventa7")
            ElseIf mytabley.Fields("pventa8") <> "" Then
                objWorksheet.Cells(v, h + 6) = "" & mytabley.Fields("pventa8")
                Punitario = "" & mytabley.Fields("pventa8")
            ElseIf mytabley.Fields("pventa9") <> "" Then
                objWorksheet.Cells(v, h + 6) = "" & mytabley.Fields("pventa9")
                Punitario = "" & mytabley.Fields("pventa9")
            ElseIf mytabley.Fields("pventa10") <> "" Then
                objWorksheet.Cells(v, h + 6) = "" & mytabley.Fields("pventa10")
                Punitario = "" & mytabley.Fields("pventa10")

            End If

            'para el saldo inicial
            'my_saldoInicial = my_saldoInicial + mytabley.Fields("saldoinicial")

            objWorksheet.Cells(v, h + 7) = my_struc_saldo_actual(j).cantidad
            my_saldo = my_saldo + my_struc_saldo_actual(j).cantidad
            objWorksheet.Cells(v, h + 8) = Round(mytabley.Fields("SUBTOTAL"), 0)
            my_total = my_total + mytabley.Fields("SUBTOTAL")
            'objWorksheet.Cells(v, h + 9) = my_struc_saldo_actual(j).fechavence
            v = v + 1
            mytabley.MoveNext

        End If

        mytabley.Close
    Next j

    'aqui los sub-totales
    objWorksheet.Cells(v, h + 6).Font.bold = True
    objWorksheet.Cells(v, h + 6).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 6) = "TOTALES"

    objWorksheet.Cells(v, h + 7).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 7) = "" & Format(my_saldo, "0.00") 'cantidad
   
    objWorksheet.Cells(v, h + 8).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 8) = "" & Round(Format(my_total, "0.00"), 0) 'cantidad
 
Detalle_saldo_actual:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

    Exit Sub

End Sub

Public Sub sele_Registro_Ventas(my_tipo As String, _
                                my_local As String, _
                                my_serie As String, _
                                my_vdetalle As String, _
                                my_numero As String, _
                                my_codigo As String, _
                                my_nombre As String, _
                                my_fechai As String, _
                                my_fechaf As String, _
                                my_moneda As String, _
                                my_vendedor As String, _
                                my_transporte As String, _
                                my_fpago As String, _
                                my_bodega As String, _
                                my_estado As String, _
                                my_grupos As String, _
                                my_cajero As String, _
                                my_consolidado As String, _
                                my_caja As String, _
                                my_turno As String, _
                                my_tipoimp As String, _
                                my_servicio As String, _
                                acu As String, _
                                salida As Boolean)
                  
    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    v = 11
    h = 0
    sdx = 0
    sdx1 = 0
    sdx2 = 0
    suma1 = 0
    suma2 = 0
    suma6 = 0
    ssuma6 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    ssuma4 = 0
    ssuma5 = 0
    suma7 = 0
    ssuma7 = 0
    suma8 = 0
    ssuma8 = 0
    suma9 = 0
    ssuma9 = 0

    On Error GoTo sele_Registro_Ventas

    'Set objWorksheet = objWorkBook.Worksheets(1)

    If Len(my_fechai) <> 10 Then Exit Sub
    If Len(my_fechaf) <> 10 Then Exit Sub
    If Not IsDate(my_fechai) Then Exit Sub
    If Not IsDate(my_fechaf) Then Exit Sub
    If my_consolidado = "S" Then
        If repdocrv.Combo1 <> "TipoDocumento" Then
            MsgBox "Grupo debe estar en TipoDocumento", 48, "Aviso"
            Exit Sub

        End If

    End If

    mysql = "select * from " & cgusuario & " where " & Chr$(10)

    If repdocrv.Check1.Value = 0 Or repdocrv.Check1.Value = 2 Then
        mysql = mysql & "  fecha>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & " and fecha<='" & Format(my_fechaf, "YYYYMMDD") & "' " & Chr$(10)

    End If

    If repdocrv.Check1.Value = 1 Then
        mysql = mysql & "  fechasunat>='" & Format(my_fechai, "YYYYMMDD") & "'" & Chr$(10)
        mysql = mysql & " and fechasunat<='" & Format(my_fechaf, "YYYYMMDD") & "' " & Chr$(10)

    End If

    If my_tipo <> "%" Then
        mysql = mysql & " and tipo like '" & extra_loquesea(my_tipo) & "'" & Chr$(10)

    End If

    If my_local <> "%" Then
        mysql = mysql & "and local like '" & extra_loquesea(my_local) & "'" & Chr$(10)

    End If

    If my_serie <> "%" Then
        mysql = mysql & " and serie like '" & my_serie & "'" & Chr$(10)

    End If

    If my_numero <> "%" Then
        mysql = mysql & " and numero like '" & my_numero & "'" & Chr$(10)

    End If

    If my_codigo <> "%" Then
        mysql = mysql & "and codigo like '" & my_codigo & "'" & Chr$(10)

    End If

    If my_nombre <> "%" Then
        mysql = mysql & " and nombre like '" & my_nombre & "'" & Chr$(10)

    End If

    If my_servicio <> "%" Then
        mysql = mysql & " and  servicio='" & extra_loquesea(my_servicio) & "'" & Chr$(10)

    End If

    If my_tipoimp = "GRAVADO" Then
        mysql = mysql & "and impuesto>0 " & Chr$(10)

    End If

    If my_tipoimp = "SERVICIO" Then
        mysql = mysql & " and servicioco>0 " & Chr$(10)

    End If

    If my_tipoimp = "EXONERADO" Then
        mysql = mysql & " and gravado>0 " & Chr$(10)

    End If

    If my_tipoimp = "IVAP" Then
        mysql = mysql & "and tivap>0 " & Chr$(10)

    End If

    If my_tipoimp = "ISC" Then
        mysql = mysql & " and tisc>0 " & Chr$(10)

    End If

    If my_tipoimp = "PERCEPCION" Then
        mysql = mysql & " and PERCEPCION>0 " & Chr$(10)

    End If

    If my_vendedor <> "%" Then
        mysql = mysql & "and vendedor like '" & extra_loquesea(my_vendedor) & "'" & Chr$(10)

    End If

    If my_transporte <> "%" Then
        mysql = mysql & " and transporte like '" & my_transporte & "'" & Chr$(10)

    End If

    If my_fpago <> "%" Then
        mysql = mysql & " and fpago like '" & my_fpago & "'" & Chr$(10)

    End If

    If my_bodega <> "%" Then
        mysql = mysql & "and bodega like '" & extra_loquesea(my_bodega) & "'" & Chr$(10)

    End If

    If my_cajero <> "%" Then
        mysql = mysql & " and usuario like '" & extra_loquesea(my_cajero) & "'" & Chr$(10)

    End If

    If my_caja <> "%" Then
        mysql = mysql & " and caja like '" & extra_loquesea(my_caja) & "'" & Chr$(10)

    End If

    If my_turno <> "%" Then
        mysql = mysql & " and turno like '" & extra_loquesea(my_turno) & "'" & Chr$(10)

    End If

    If acu = "C" Then
        If repdocrv.Check2.Value = 0 Or repdocrv.Check1.Value = 2 Then
            mysql = mysql & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O')" & Chr$(10)

        End If

        If repdocrv.Check2.Value = 1 Then
            mysql = mysql & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='N' or acu='O')" & Chr$(10)

        End If
   
    End If

    If acu = "V" Then
        If repdocrv.Check2.Value = 0 Or repdocrv.Check1.Value = 2 Then
            mysql = mysql & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F')" & Chr$(10)

        End If

        If repdocrv.Check2.Value = 1 Then
            mysql = mysql & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D'  or acu='E' or acu='F')" & Chr$(10)

        End If
   
    End If

    If my_estado <> "%" Then
        mysql = mysql & " and estado='" & estado & "'" & Chr$(10)

    End If

    If repdocrv.Combo1 = "TipoDocumento" Then
        mysql = mysql & "order by tipo,fecha" & Chr$(10)

    End If

    If repdocrv.Combo1 = "Codigo" Then
        mysql = mysql & "order by Codigo,fecha" & Chr$(10)

    End If

    If repdocrv.Combo1 = "Vendedor" Then
        mysql = mysql & "order by vendedor,fecha" & Chr$(10)

    End If

    If repdocrv.Combo1 = "Zona" Then
        mysql = mysql & "order by Zona,fecha" & Chr$(10)

    End If

    If repdocrv.Combo1 = "Fecha" Then
        mysql = mysql & "order by Fecha,Local,tipo,serie,str(numero)" & Chr$(10)

    End If

    mytablex.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF Then
        salida = False
        Exit Sub
    Else
        salida = True
        mytablex.MoveFirst

        Do Until mytablex.EOF
            Set objWorksheet = objWorkBook.Worksheets(1)

            'If v = 13 Then
            If mytablex.Fields("MONEDA") = "S" Then
                objWorksheet.Cells(8, 12).Font.bold = True
                objWorksheet.Cells(8, 12).Font.Size = 5
                objWorksheet.Cells(8, 12).Select
                Selection.ColumnWidth = 9
                  
                objWorksheet.Cells(8, 12) = "Moneda"
                objWorksheet.Cells(8, 13).Font.Size = 5
                objWorksheet.Cells(8, 13) = "Soles"
            Else
                objWorksheet.Cells(8, 12).Font.bold = True
                objWorksheet.Cells(8, 12).Font.Size = 5
                objWorksheet.Cells(8, 12).Select
                Selection.ColumnWidth = 9
                objWorksheet.Cells(8, 12) = "Moneda"
                objWorksheet.Cells(8, 13) = "Dolares"

            End If

            'End If
            objWorksheet.Cells(v, h + 1).Font.Size = 7 'aqui tamaño letra pll
            objWorksheet.Cells(v, h + 1) = "'" & Format(mytablex.Fields("fecha"), "DD/MM/YYYY")
            objWorksheet.Cells(v, h + 2).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 2) = mytablex.Fields("local")
            objWorksheet.Cells(v, h + 3).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 3) = mytablex.Fields("tipo")
            objWorksheet.Cells(v, h + 4).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 4) = mytablex.Fields("serie")
            objWorksheet.Cells(v, h + 5).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 5) = mytablex.Fields("Numero")
            objWorksheet.Cells(v, h + 6).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 6) = mytablex.Fields("Codigo")
            objWorksheet.Cells(v, h + 7).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 7) = mytablex.Fields("Nombre")

            If "" & mytablex.Fields("estado") = "1" Then
                objWorksheet.Cells(v, h + 8).Font.Size = 4 'aqui tamaño letra
                objWorksheet.Cells(v, h + 8) = "ANULADO"
                objWorksheet.Range(Cells(v, h + 8), Cells(v, 17)).Interior.color = RGB(255, 164, 32)  'Color anaranjado
                'GoTo amiga11
            Else
                objWorksheet.Cells(v, h + 8).Font.Size = 4 'aqui tamaño letra
                objWorksheet.Cells(v, h + 8) = "ACTIVADO"

            End If

            If mytablex.Fields("moneda") = "D" Then
                objWorksheet.Cells(v, h + 9).Font.Size = 7 'aqui tamaño letra
                'objWorksheet.Cells(v, h + 9) = Format(Val("" & mytablex.Fields("subtotal")) * _
                 Val(xparidad) - Val("" & mytablex.Fields("gravado")) _
                 * Val(xparidad), "0.00")

                'inicio 10/11/2017 pll
                If mytablex.Fields("tipo1") = "NCV" Then
                    objWorksheet.Cells(v, h + 9) = "'" & Format(Round(Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("gravado")), 2), "0.00")
                                        
                    objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 10) = "-" & Format(mytablex.Fields("gravado") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 11) = "-" & Format(mytablex.Fields("tisc") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letra
                    'objWorksheet.Cells(v, h + 12) = Format(mytablex.Fields("impuesto") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 12) = "-" & Format(mytablex.Fields("impuesto"))
                    objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra
                    'objWorksheet.Cells(v, h + 13) = Format(mytablex.Fields("total") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 13) = "-" & Format(mytablex.Fields("total"))
                    objWorksheet.Cells(v, h + 14).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 14) = "-" & Format(mytablex.Fields("tivap") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 15).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 15) = "-" & Format(mytablex.Fields("percepcion") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 16).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 16) = "-" & Format(mytablex.Fields("servicioco") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 17).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 17) = "-" & Format(mytablex.Fields("tdetra") * Val(xparidad), "0.00")
       
                Else
                    objWorksheet.Cells(v, h + 9) = "" & Format(Round(Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("gravado")), 2), "0.00")
                                        
                    objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 10) = "'" & Format(Round(Val("" & mytablex.Fields("gravado") * Val(xparidad)), 2), "0.00")
                    objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 11) = "'" & Format(Round(Val("" & mytablex.Fields("tisc") * Val(xparidad)), 2), "0.00")
                    objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letra
                    ''objWorksheet.Cells(v, h + 12) = Format(mytablex.Fields("impuesto") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 12) = "'" & Format(Round(Val("" & mytablex.Fields("impuesto")), 2), "0.00")
                    objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra
                    ''objWorksheet.Cells(v, h + 13) = Format(mytablex.Fields("total") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 13) = "'" & Format(Round(Val("" & mytablex.Fields("total")), 2), "0.00")
                    objWorksheet.Cells(v, h + 14).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 14) = "'" & Format(Round(Val("" & mytablex.Fields("tivap") * Val(xparidad)), 2), "0.00")
                    objWorksheet.Cells(v, h + 15).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 15) = "'" & Format(Round(Val("" & mytablex.Fields("percepcion") * Val(xparidad)), 2), "0.00")
                    objWorksheet.Cells(v, h + 16).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 16) = "'" & Format(Round(Val("" & mytablex.Fields("servicioco") * Val(xparidad)), 2), "0.00")
                    objWorksheet.Cells(v, h + 17).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 17) = "'" & Format(Round(Val("" & mytablex.Fields("tdetra") * Val(xparidad)), 2), "0.00")

                End If
       
                If mytablex.Fields("acu") = "E" Then
                    objWorksheet.Range(Cells(v, h + 9), Cells(v, 17)).Interior.color = RGB(255, 164, 32)  'Color anaranjado
          
                    objWorksheet.Cells(v, h + 9) = "-" & Format(Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("gravado")), "0.00")
                    objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 10) = "-" & Format(mytablex.Fields("gravado") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 11) = "-" & Format(mytablex.Fields("tisc") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letra
                    'objWorksheet.Cells(v, h + 12) = Format(mytablex.Fields("impuesto") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 12) = "-" & Format(mytablex.Fields("impuesto"))
                    objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra
                    'objWorksheet.Cells(v, h + 13) = Format(mytablex.Fields("total") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 13) = "-" & Format(mytablex.Fields("total"))
                    objWorksheet.Cells(v, h + 14).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 14) = "-" & Format(mytablex.Fields("tivap") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 15).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 15) = "-" & Format(mytablex.Fields("percepcion") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 16).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 16) = "-" & Format(mytablex.Fields("servicioco") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 17).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 17) = "-" & Format(mytablex.Fields("tdetra") * Val(xparidad), "0.00")
                Else

                    If mytablex.Fields("tipo1") = "NCV" Then
                        objWorksheet.Cells(v, h + 9) = "-" & Format(Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("gravado")), "0.00")
                        objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
                        objWorksheet.Cells(v, h + 10) = "-" & Format(mytablex.Fields("gravado") * Val(xparidad), "0.00")
                        objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
                        objWorksheet.Cells(v, h + 11) = "-" & Format(mytablex.Fields("tisc") * Val(xparidad), "0.00")
                        objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letra
                        'objWorksheet.Cells(v, h + 12) = Format(mytablex.Fields("impuesto") * Val(xparidad), "0.00")
                        objWorksheet.Cells(v, h + 12) = "-" & Format(mytablex.Fields("impuesto"))
                        objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra
                        'objWorksheet.Cells(v, h + 13) = Format(mytablex.Fields("total") * Val(xparidad), "0.00")
                        objWorksheet.Cells(v, h + 13) = "-" & Format(mytablex.Fields("total"))
                        objWorksheet.Cells(v, h + 14).Font.Size = 7 'aqui tamaño letra
                        objWorksheet.Cells(v, h + 14) = "-" & Format(mytablex.Fields("tivap") * Val(xparidad), "0.00")
                        objWorksheet.Cells(v, h + 15).Font.Size = 7 'aqui tamaño letra
                        objWorksheet.Cells(v, h + 15) = "-" & Format(mytablex.Fields("percepcion") * Val(xparidad), "0.00")
                        objWorksheet.Cells(v, h + 16).Font.Size = 7 'aqui tamaño letra
                        objWorksheet.Cells(v, h + 16) = "-" & Format(mytablex.Fields("servicioco") * Val(xparidad), "0.00")
                        objWorksheet.Cells(v, h + 17).Font.Size = 7 'aqui tamaño letra
                        objWorksheet.Cells(v, h + 17) = "-" & Format(mytablex.Fields("tdetra") * Val(xparidad), "0.00")

                    End If

                End If

                'fin 10/11/2017 pll
                'objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
                'objWorksheet.Cells(v, h + 10) = Format(mytablex.Fields("gravado") * Val(xparidad), "0.00")
                'objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
                'objWorksheet.Cells(v, h + 11) = Format(mytablex.Fields("tisc") * Val(xparidad), "0.00")
                'objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letra
                'objWorksheet.Cells(v, h + 12) = Format(mytablex.Fields("impuesto") * Val(xparidad), "0.00")
                'objWorksheet.Cells(v, h + 12) = Format(mytablex.Fields("impuesto"))
                'objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra
                'objWorksheet.Cells(v, h + 13) = Format(mytablex.Fields("total") * Val(xparidad), "0.00")
                'objWorksheet.Cells(v, h + 13) = Format(mytablex.Fields("total"))
                'objWorksheet.Cells(v, h + 14).Font.Size = 7 'aqui tamaño letra
                'objWorksheet.Cells(v, h + 14) = Format(mytablex.Fields("tivap") * Val(xparidad), "0.00")
                'objWorksheet.Cells(v, h + 15).Font.Size = 7 'aqui tamaño letra
                'objWorksheet.Cells(v, h + 15) = Format(mytablex.Fields("percepcion") * Val(xparidad), "0.00")
                'objWorksheet.Cells(v, h + 16).Font.Size = 7 'aqui tamaño letra
                'objWorksheet.Cells(v, h + 16) = Format(mytablex.Fields("servicioco") * Val(xparidad), "0.00")
                'objWorksheet.Cells(v, h + 17).Font.Size = 7 'aqui tamaño letra
                'objWorksheet.Cells(v, h + 17) = Format(mytablex.Fields("tdetra") * Val(xparidad), "0.00")
            End If

            If mytablex.Fields("moneda") = "S" Then
                objWorksheet.Cells(v, h + 9).Font.Size = 7 'aqui tamaño letra

                'inicio 10/11/2017 pll
                If mytablex.Fields("tipo1") = "NCV" Then
                    objWorksheet.Cells(v, h + 9) = "'" & Format(Round(Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("gravado")), 2), "0.00")
                    objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 10) = "-" & Format(mytablex.Fields("gravado") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 11) = "-" & Format(mytablex.Fields("tisc") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letra
                    'objWorksheet.Cells(v, h + 12) = Format(mytablex.Fields("impuesto") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 12) = "-" & Format(mytablex.Fields("impuesto"))
                    objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra
                    'objWorksheet.Cells(v, h + 13) = Format(mytablex.Fields("total") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 13) = "-" & Format(mytablex.Fields("total"))
                    objWorksheet.Cells(v, h + 14).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 14) = "-" & Format(mytablex.Fields("tivap") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 15).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 15) = "-" & Format(mytablex.Fields("percepcion") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 16).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 16) = "-" & Format(mytablex.Fields("servicioco") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 17).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 17) = "-" & Format(mytablex.Fields("tdetra") * Val(xparidad), "0.00")
                Else
                    objWorksheet.Cells(v, h + 9) = "" & Format(Round(Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("gravado")), 2), "0.00")
                    objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 10) = Format(mytablex.Fields("gravado") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 11) = Format(mytablex.Fields("tisc") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letra
                    ''objWorksheet.Cells(v, h + 12) = Format(mytablex.Fields("impuesto") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 12) = Format(mytablex.Fields("impuesto"))
                    objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra
                    ''objWorksheet.Cells(v, h + 13) = Format(mytablex.Fields("total") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 13) = Format(mytablex.Fields("total"))
                    objWorksheet.Cells(v, h + 14).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 14) = Format(mytablex.Fields("tivap") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 15).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 15) = Format(mytablex.Fields("percepcion") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 16).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 16) = Format(mytablex.Fields("servicioco") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 17).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 17) = Format(mytablex.Fields("tdetra") * Val(xparidad), "0.00")

                End If
       
                If mytablex.Fields("acu") = "E" Then
                    objWorksheet.Cells(v, h + 9) = "'" & Format(Round(Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("gravado")), 2), "0.00")
                    objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 10) = "-" & Format(mytablex.Fields("gravado") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 11) = "-" & Format(mytablex.Fields("tisc") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letra
                    'objWorksheet.Cells(v, h + 12) = Format(mytablex.Fields("impuesto") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 12) = "-" & Format(mytablex.Fields("impuesto"))
                    objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra
                    'objWorksheet.Cells(v, h + 13) = Format(mytablex.Fields("total") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 13) = "-" & Format(mytablex.Fields("total"))
                    objWorksheet.Cells(v, h + 14).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 14) = "-" & Format(mytablex.Fields("tivap") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 15).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 15) = "-" & Format(mytablex.Fields("percepcion") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 16).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 16) = "-" & Format(mytablex.Fields("servicioco") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 17).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 17) = "-" & Format(mytablex.Fields("tdetra") * Val(xparidad), "0.00")
                Else
                    objWorksheet.Cells(v, h + 9) = "" & Format(Round(Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("gravado")), 2), "0.00")
                    objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 10) = Format(mytablex.Fields("gravado") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 11) = Format(mytablex.Fields("tisc") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letra
                    ''objWorksheet.Cells(v, h + 12) = Format(mytablex.Fields("impuesto") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 12) = Format(mytablex.Fields("impuesto"))
                    objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra
                    ''objWorksheet.Cells(v, h + 13) = Format(mytablex.Fields("total") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 13) = Format(mytablex.Fields("total"))
                    objWorksheet.Cells(v, h + 14).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 14) = Format(mytablex.Fields("tivap") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 15).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 15) = Format(mytablex.Fields("percepcion") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 16).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 16) = Format(mytablex.Fields("servicioco") * Val(xparidad), "0.00")
                    objWorksheet.Cells(v, h + 17).Font.Size = 7 'aqui tamaño letra
                    objWorksheet.Cells(v, h + 17) = Format(mytablex.Fields("tdetra") * Val(xparidad), "0.00")

                End If

                'fin 10/11/2017 pll
                'objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
                'objWorksheet.Cells(v, h + 10) = Format(mytablex.Fields("gravado"), "0.00")
                'objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
                'objWorksheet.Cells(v, h + 11) = Format(mytablex.Fields("tisc"), "0.00")
                'objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letra
                'objWorksheet.Cells(v, h + 12) = Format(mytablex.Fields("impuesto"), "0.00")
                'objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra
                'objWorksheet.Cells(v, h + 13) = Format(mytablex.Fields("total"), "0.00")
                'objWorksheet.Cells(v, h + 14).Font.Size = 7 'aqui tamaño letra
                'objWorksheet.Cells(v, h + 14) = Format(mytablex.Fields("tivap"), "0.00")
                'objWorksheet.Cells(v, h + 15).Font.Size = 7 'aqui tamaño letra
                'objWorksheet.Cells(v, h + 15) = Format(mytablex.Fields("percepcion"), "0.00")
                'objWorksheet.Cells(v, h + 16).Font.Size = 7 'aqui tamaño letra
                'objWorksheet.Cells(v, h + 16) = Format(mytablex.Fields("SERVICIOCO"), "0.00")
                'objWorksheet.Cells(v, h + 17).Font.Size = 7 'aqui tamaño letra
                'objWorksheet.Cells(v, h + 17) = Format(mytablex.Fields("tdetra"), "0.00")
            End If
    
            v = v + 1
    
            If "" & mytablex.Fields("moneda") = "S" And "" & mytablex.Fields("estado") = "2" Then
                If mytablex.Fields("acu") <> "E" And mytablex.Fields("tipo1") <> "NCV" Then
                    suma1 = suma1 + Format(Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("gravado")), "0.00")
                    suma2 = suma2 + Format(Val("" & mytablex.Fields("gravado")), "0.00")
                    suma6 = suma6 + Format(Val("" & mytablex.Fields("tivap")), "0.00")
                    ssuma6 = ssuma6 + Format(Val("" & mytablex.Fields("tivap")), "0.00")
                    suma3 = suma3 + Format(Val("" & mytablex.Fields("tisc")), "0.00")
      
                    suma4 = suma4 + Format(Val("" & mytablex.Fields("impuesto")), "0.00")
                    suma5 = suma5 + Format(Val("" & mytablex.Fields("total")), "0.00")
                    ssuma1 = ssuma1 + Format(Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("gravado")), "0.00")
                    ssuma2 = ssuma2 + Format(Val("" & mytablex.Fields("gravado")), "0.00")
                    ssuma3 = ssuma3 + Format(Val("" & mytablex.Fields("tisc")), "0.00")
                    ssuma4 = ssuma4 + Format(Val("" & mytablex.Fields("impuesto")), "0.00")
                    ssuma5 = ssuma5 + Format(Val("" & mytablex.Fields("total")), "0.00")
      
                    suma7 = suma7 + Format(Val("" & mytablex.Fields("percepcion")), "0.00")
                    ssuma7 = ssuma7 + Format(Val("" & mytablex.Fields("percepcion")), "0.00")
      
                    suma8 = suma8 + Format(Val("" & mytablex.Fields("SERVICIOCO")), "0.00")
                    ssuma8 = ssuma8 + Format(Val("" & mytablex.Fields("SERVICIOCO")), "0.00")
      
                    suma9 = suma9 + Format(Val("" & mytablex.Fields("tdetra")), "0.00")
                    ssuma9 = ssuma9 + Format(Val("" & mytablex.Fields("tdetra")), "0.00")

                End If

            End If
   
            If "" & mytablex.Fields("moneda") = "D" And "" & mytablex.Fields("estado") = "2" Then
                'suma1 = suma1 + Format(Val(Format(Val("" & mytablex.Fields("subtotal")) * Val(xparidad) - _
                 Val("" & mytablex.Fields("gravado")) * Val(xparidad), "0.00")), "0.00")

                If mytablex.Fields("acu") <> "E" And mytablex.Fields("tipo1") <> "NCV" Then
                    suma1 = suma1 + Format(Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("gravado")), "0.00")
     
                    suma2 = suma2 + Format(Val(Format(Val("" & mytablex.Fields("gravado")) * Val(xparidad), "0.00")), "0.00")
                    suma6 = suma6 + Format(Val(Format(Val("" & mytablex.Fields("tivap")) * Val(xparidad), "0.00")), "0.00")
                    ssuma6 = ssuma6 + Format(Val(Format(Val("" & mytablex.Fields("tivap")) * Val(xparidad), "0.00")), "0.00")
      
                    suma3 = suma3 + Format(Val(Format(Val("" & mytablex.Fields("tisc")) * Val(xparidad), "0.00")), "0.00")
                    ssuma3 = ssuma3 + Format(Val(Format(Val("" & mytablex.Fields("tisc")) * Val(xparidad), "0.00")), "0.00")
      
                    'suma4 = suma4 + Format(Val(Format(Val("" & mytablex.Fields("impuesto")) * Val(xparidad), "0.00")), "0.00")
                    suma4 = suma4 + Format(Val(Format(Val("" & mytablex.Fields("impuesto")), "0.00")), "0.00")
                    'suma5 = suma5 + Format(Val(Format(Val("" & mytablex.Fields("total")) * Val(xparidad), "0.00")), "0.00")
                    suma5 = suma5 + Format(Val(Format(Val("" & mytablex.Fields("total")), "0.00")), "0.00")
                    'ssuma1 = ssuma1 + Format(Val(Format(Val("" & mytablex.Fields("subtotal")) * Val(xparidad) - _
                     Val("" & mytablex.Fields("gravado")) * Val(xparidad), "0.00")), "0.00")
                    ssuma1 = ssuma1 + Format(Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("gravado")), "0.00")
     
                    ssuma2 = ssuma2 + Format(Val(Format(Val("" & mytablex.Fields("gravado")) * Val(xparidad), "0.00")), "0.00")
                    'ssuma4 = ssuma4 + Format(Val(Format(Val("" & mytablex.Fields("impuesto")) * Val(xparidad), "0.00")), "0.00")
                    ssuma4 = ssuma4 + suma4 'Format(Val(Format(Val("" & mytablex.Fields("impuesto")) * Val(xparidad), "0.00")), "0.00")
                    'ssuma5 = ssuma5 + Format(Val(Format(Val("" & mytablex.Fields("total")) * Val(xparidad), "0.00")), "0.00")
                    ssuma5 = ssuma5 + Format(Val(Format(Val("" & mytablex.Fields("total")), "0.00")), "0.00")
    
                    suma7 = suma7 + Format(Val(Format(Val("" & mytablex.Fields("percepcion")) * Val(xparidad), "0.00")), "0.00")
                    ssuma7 = ssuma7 + Format(Val(Format(Val("" & mytablex.Fields("percepcion")) * Val(xparidad), "0.00")), "0.00")
                    suma8 = suma8 + Format(Val(Format(Val("" & mytablex.Fields("servicioco")) * Val(xparidad), "0.00")), "0.00")
                    ssuma8 = ssuma8 + Format(Val(Format(Val("" & mytablex.Fields("servicioco")) * Val(xparidad), "0.00")), "0.00")
      
                    suma9 = suma9 + Format(Val(Format(Val("" & mytablex.Fields("tdetra")) * Val(xparidad), "0.00")), "0.00")
                    ssuma9 = ssuma9 + Format(Val(Format(Val("" & mytablex.Fields("tdetra")) * Val(xparidad), "0.00")), "0.00")

                End If

            End If

            mytablex.MoveNext

        Loop

        mytablex.Close

    End If

    '*****************************
    'aqui los sub-totales
    'h = 3
    'objWorksheet.Cells(v, h + 5).Font.Bold = True
    'objWorksheet.Cells(v, h + 5).Interior.color = RGB(215, 215, 0)  'resaltador
    'objWorksheet.Cells(v, h + 5) = "Sub-TOTALES"
   
    'objWorksheet.Cells(v, h + 6).Interior.color = RGB(215, 215, 0)  'resaltador
    'objWorksheet.Cells(v, h + 6) = "" & Format(suma1, "0.00") 'cantidad
   
    'objWorksheet.Cells(v, h + 7).Interior.color = RGB(215, 215, 0)  'resaltador
    'objWorksheet.Cells(v, h + 7) = "" & Format(suma2, "0.00") 'precio unitario
   
    'objWorksheet.Cells(v, h + 8).Interior.color = RGB(215, 215, 0)  'resaltador
    'objWorksheet.Cells(v, h + 8) = "" & Format(suma3, "0.00") ''vtaxproc.total
   
    'objWorksheet.Cells(v, h + 9).Interior.color = RGB(215, 215, 0)  'resaltador
    'objWorksheet.Cells(v, h + 9) = "" & Format(suma4, "0.00") 'costo unitario
   
    'objWorksheet.Cells(v, h + 10).Interior.color = RGB(215, 215, 0)  'resaltador
    'objWorksheet.Cells(v, h + 10) = "" & Format(suma5, "0.00") 'costo total
   
    'objWorksheet.Cells(v, h + 11).Interior.color = RGB(215, 215, 0)  'resaltador
    'objWorksheet.Cells(v, h + 11) = "" & Format(suma6, "0.00") 'ganancia
    h = 3
    v = v + 1
    'aqui los totales
    'provemos anteriormente tambien estaba asi
    objWorksheet.Cells(v, h + 5).Font.bold = True
    objWorksheet.Cells(v, h + 5).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 5) = "TOTALES"
   
    objWorksheet.Cells(v, h + 6).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 6) = "'" & Format(Round(Val(ssuma1), 2), "0.00") 'cantidad

    objWorksheet.Cells(v, h + 7).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 7) = "" & Format(ssuma2, "0.00") 'precio unitario
                                       
    objWorksheet.Cells(v, h + 8).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 8) = "" & Format(ssuma3, "0.00") ''vtaxproc.total
   
    objWorksheet.Cells(v, h + 9).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 9) = "'" & Format(Round(Val(suma4), 2), "0.00") 'costo unitario
   
    objWorksheet.Cells(v, h + 10).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 10) = "'" & Format(Round(Val(ssuma5), 2), "0.00")  'costo total
   
    objWorksheet.Cells(v, h + 11).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 11) = "" & Format(ssuma6, "0.00")  'ganancia
    Exit Sub
sele_Registro_Ventas:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

End Sub

Public Sub titulo_frmChart2(fechai As String, fechaf As String)

    On Error GoTo titulo_frmChart

    Set objWorksheet = objWorkBook.Worksheets(1)

    objWorksheet.Cells(11, 1).Font.bold = True
    objWorksheet.Cells(11, 1).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(11, 1) = "Desde"
    objWorksheet.Cells(11, 2).Select
    Selection.ColumnWidth = 9
    objWorksheet.Cells(11, 2).Font.Size = 4 'aqui tamaño letra
    objWorksheet.Cells(11, 2) = fechai
 
    objWorksheet.Cells(11, 3).Font.bold = True
    objWorksheet.Cells(11, 3).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(11, 3) = "Al"
    objWorksheet.Cells(11, 4).Select
    Selection.ColumnWidth = 9
    objWorksheet.Cells(11, 4).Font.Size = 4 'aqui tamaño letra
    objWorksheet.Cells(11, 4) = fechaf
 
    'Aqui los margenes cuadrados
    objWorksheet.Range(Cells(12, 1), Cells(12, 3)).Borders.LineStyle = xlContinuous 'Codigo
    objWorksheet.Range(Cells(12, 1), Cells(12, 3)).Interior.color = RGB(215, 215, 215) 'Codigo

    'fecha a imprimir
    objWorksheet.Cells(10, 1).Font.bold = True
    objWorksheet.Cells(10, 1).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 1) = "FECHA HOY :"
    objWorksheet.Cells(10, 2).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(10, 2) = Format(Now, "dd/mm/yyyy") & "- HORA " + Format(Now, "HH:MM:SS")
 
    objWorksheet.Cells(12, 1).Font.bold = True
    objWorksheet.Cells(12, 1).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(12, 1) = "mes"
    objWorksheet.Range("A10").ColumnWidth = 8
  
    objWorksheet.Cells(12, 2).Font.bold = True
    objWorksheet.Cells(12, 2).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(12, 2) = "Moneda"
    objWorksheet.Range("B10").ColumnWidth = 8
  
    objWorksheet.Cells(12, 3).Font.bold = True
    objWorksheet.Cells(12, 2).Font.Size = 7 'aqui tamaño letra
    objWorksheet.Cells(12, 3) = "Total"
    objWorksheet.Range("C10").ColumnWidth = 8
  
    Exit Sub
titulo_frmChart:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

    Exit Sub

End Sub

Public Sub carga_impresion_total(my_struc_cotizacion_total_excel() As struc_cotizacion_total_excel, _
                                 D As Integer, _
                                 my_moneda As String, _
                                 my_transporte As String)
           
    v = 13
    h = 0
    sdx1 = 0
    sdx = 0
    subtotal = 0
    igv = 0

    On Error GoTo carga_impresion_total

    For I = 0 To D - 1
        sdx = sdx + Val(my_struc_cotizacion_total_excel(I).cantidad)
        sdx1 = sdx1 + Val(my_struc_cotizacion_total_excel(I).total)
        subtotal = subtotal + Val(my_struc_cotizacion_total_excel(I).subtotal)
        igv = igv + Val(my_struc_cotizacion_total_excel(I).igv)
  
        objWorksheet.Cells(v, h + 1) = my_struc_cotizacion_total_excel(I).producto
        objWorksheet.Cells(v, h + 2) = "'" & my_struc_cotizacion_total_excel(I).descripcion
        objWorksheet.Cells(v, h + 3) = my_struc_cotizacion_total_excel(I).unidad
        objWorksheet.Cells(v, h + 4) = my_struc_cotizacion_total_excel(I).factor
        objWorksheet.Cells(v, h + 5) = my_struc_cotizacion_total_excel(I).cantidad
        objWorksheet.Cells(v, h + 6) = my_struc_cotizacion_total_excel(I).precio
        objWorksheet.Cells(v, h + 7) = my_struc_cotizacion_total_excel(I).total
  
        v = v + 1
    Next I

    'vemos en donde impacta 16/08/2017 pll
    If Mid(my_moneda, 1, 1) = "S" Then
        '  If my_moneda = "S" Then
        'inicio 16/08/2017 pll subtotal
        objWorksheet.Cells(v, h + 4).Font.bold = True
        objWorksheet.Cells(v, h + 4).Interior.color = RGB(215, 215, 0)  'resaltador
        objWorksheet.Cells(v, h + 4) = "Sub-TOTAL"
        
        objWorksheet.Cells(v, h + 7).Interior.color = RGB(215, 215, 0)  'resaltador
        objWorksheet.Cells(v, h + 7) = "" & Format(subtotal, "0.00") 'total subtotal
        'fin 16/08/2017 pll subtotal

        'inicio 17/08/2017 pll igv
        objWorksheet.Cells(v + 1, h + 4).Font.bold = True
        objWorksheet.Cells(v + 1, h + 4).Interior.color = RGB(215, 215, 0) 'resaltador
        objWorksheet.Cells(v + 1, h + 4) = "IGV"
        
        objWorksheet.Cells(v + 1, h + 7).Interior.color = RGB(215, 215, 0) 'resaltador
        objWorksheet.Cells(v + 1, h + 7) = "" & Format(igv, "0.00") 'total subtotal
        'fin 17/08/2017 pll igv

        objWorksheet.Cells(v + 2, h + 4).Font.bold = True
        objWorksheet.Cells(v + 2, h + 4).Interior.color = RGB(215, 215, 0) 'resaltador
        objWorksheet.Cells(v + 2, h + 4) = "TOTAL SOLES"
        
        objWorksheet.Cells(v + 2, h + 5).Interior.color = RGB(215, 215, 0) 'resaltador
        objWorksheet.Cells(v + 2, h + 5) = "" & Format(sdx, "0.00") 'total cantidad
   
        objWorksheet.Cells(v + 2, h + 7).Interior.color = RGB(215, 215, 0) 'resaltador
        objWorksheet.Cells(v + 2, h + 7) = "" & Format(sdx1, "0.00") 'total total
      
    Else
        'inicio 16/08/2017 pll subtotal
        objWorksheet.Cells(v, h + 4).Font.bold = True
        objWorksheet.Cells(v, h + 4).Interior.color = RGB(215, 215, 0)  'resaltador
        objWorksheet.Cells(v, h + 4) = "Sub-TOTAL"
        
        objWorksheet.Cells(v, h + 7).Interior.color = RGB(215, 215, 0)  'resaltador
        objWorksheet.Cells(v, h + 7) = "" & Format(subtotal, "0.00") 'total subtotal
        'fin 16/08/2017 pll subtotal

        'inicio 17/08/2017 pll igv
        objWorksheet.Cells(v + 1, h + 4).Font.bold = True
        objWorksheet.Cells(v + 1, h + 4).Interior.color = RGB(215, 215, 0) 'resaltador
        objWorksheet.Cells(v + 1, h + 4) = "IGV"
        
        objWorksheet.Cells(v + 1, h + 7).Interior.color = RGB(215, 215, 0) 'resaltador
        objWorksheet.Cells(v + 1, h + 7) = "" & Format(igv, "0.00") 'total subtotal
        'fin 17/08/2017 pll igv

        objWorksheet.Cells(v + 2, h + 4).Font.bold = True
        objWorksheet.Cells(v + 2, h + 4).Interior.color = RGB(215, 215, 0) 'resaltador
        objWorksheet.Cells(v + 2, h + 4) = "TOTAL DOLARES"
        
        objWorksheet.Cells(v + 2, h + 5).Interior.color = RGB(215, 215, 0) 'resaltador
        objWorksheet.Cells(v + 2, h + 5) = "" & Format(sdx, "0.00") 'total cantidad
   
        objWorksheet.Cells(v + 2, h + 7).Interior.color = RGB(215, 215, 0) 'resaltador
        objWorksheet.Cells(v + 2, h + 7) = "" & Format(sdx1, "0.00") 'total total

    End If
 
    Exit Sub
 
carga_impresion_total:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

End Sub

Public Sub carga_solo_documentos(my_struc_solo_documentos() As struc_solo_documentos, _
                                 k As Integer, _
                                 my_moneda As String)
    v = 13
    h = 0
    my_total = 0

    On Error GoTo carga_solo_documentos

    For I = 0 To k - 1
        my_total = my_total + Val(my_struc_solo_documentos(I).total)
 
        objWorksheet.Cells(v, h + 1) = my_struc_solo_documentos(I).codigo
        objWorksheet.Cells(v, h + 2) = "'" & my_struc_solo_documentos(I).nombre
        objWorksheet.Cells(v, h + 3) = my_struc_solo_documentos(I).local
        objWorksheet.Cells(v, h + 4) = my_struc_solo_documentos(I).estado
        objWorksheet.Cells(v, h + 5) = my_struc_solo_documentos(I).tipo
        objWorksheet.Cells(v, h + 6) = my_struc_solo_documentos(I).serie
        objWorksheet.Cells(v, h + 7) = my_struc_solo_documentos(I).Numero
        objWorksheet.Cells(v, h + 8) = my_struc_solo_documentos(I).fecha
        objWorksheet.Cells(v, h + 9) = my_struc_solo_documentos(I).total
        v = v + 1
    Next I

    If my_moneda = "Soles" Then
        objWorksheet.Cells(v, h + 4).Font.bold = True
        objWorksheet.Cells(v, h + 4).Interior.color = RGB(215, 215, 0)  'resaltador
        objWorksheet.Cells(v, h + 4) = "TOTAL SOLES"
   
        '      objWorksheet.Cells(v, h + 5).Interior.color = RGB(215, 215, 0)  'resaltador
        '      objWorksheet.Cells(v, h + 5) = "" & Format(my_total, "0.00")
   
        objWorksheet.Cells(v, h + 9).Interior.color = RGB(215, 215, 0)  'resaltador
        objWorksheet.Cells(v, h + 9) = "" & Format(my_total, "0.00")
    Else
        objWorksheet.Cells(v, h + 4).Font.bold = True
        objWorksheet.Cells(v, h + 4).Interior.color = RGB(215, 215, 0)  'resaltador
        objWorksheet.Cells(v, h + 4) = "TOTAL DOLARES"
   
        '    objWorksheet.Cells(v, h + 5).Interior.color = RGB(215, 215, 0)  'resaltador
        '    objWorksheet.Cells(v, h + 5) = "" & Format(my_total, "0.00")
   
        objWorksheet.Cells(v, h + 9).Interior.color = RGB(215, 215, 0)  'resaltador
        objWorksheet.Cells(v, h + 9) = "" & Format(my_total, "0.00")

    End If
 
    Exit Sub
 
carga_solo_documentos:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

End Sub

Public Sub cerra_excel(my_file As String)

    Dim myPDF As PdfDistiller

    'inicio 09/08/2017 pll eso proteje con escritura
    objWorkBook.Protect Structure:=True, Windows:=True
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True

    If Mid(my_file, 1, 8) = "VGrafico" Then
        Worksheets(1).Protect UserInterfaceOnly:=True

    End If

    'fin 09/08720917 pll

    If Dir("C:\Reporte") = "" Then
        Call crea_directorio_excel

    End If

    objWorkBook.SaveAs "C:\Reporte\" & my_file & ".xls"

    Set objExcel = New Excel.Application
    Set xlibro = objExcel.Workbooks.Open("C:\Reporte\" & my_file & ".xls", vbMaximizedFocus)
    objExcel.Visible = True

    'objWorksheet.PageSetup.Orientation = vbHorizontal
    If Mid(my_file, 1, 8) <> "VGrafico" Then
        ActiveSheet.PageSetup.Orientation = xlLandscape
        'wsXL.PageSetup.Orientation = vbHorizontal
        'inicio conver ps
        objWorksheet.PrintOut , , , , "Adobe PDF", True, , "C:\Reporte\test.PS"

        Set myPDF = New PdfDistiller
        myPDF.FileToPDF "C:\Reporte\test.ps", "C:\Reporte\" & my_file & ".pdf", ""

        'fin
    End If

    'objWorkBook.Close

    Set objWorkBook = Nothing
    Set objWorksheet = Nothing
    oXL.Quit
    Set oXL = Nothing
    'Delete the old PS file
    Kill "C:\Reporte\test.ps"

End Sub

''''''16/08/2017 pll
Public Function control_trasporte(my_local As String, _
                                  my_tipo As String, _
                                  my_serie As String, _
                                  my_numero As String, _
                                  my_transporte As String, _
                                  salida As Boolean, _
                                  my_struc_Etransporte() As struc_Etransporte, _
                                  k As Integer)
           
    Dim mysql    As String

    Dim mytablex As New ADODB.Recordset

    ReDim my_struc_Etransporte(0)

    mysql = "SELECT distinct " & Chr$(10)
    mysql = mysql & "f.partida,f.destino,f.fecha," & Chr$(10)
    mysql = mysql & "t.codigo as ruc,t.nombre as nombreT,t.nombrec as nombreC," & Chr$(10)
    mysql = mysql & "t.placa," & Chr$(10)
    mysql = mysql & "t.licencia," & Chr$(10)
    mysql = mysql & "t.marca," & Chr$(10)
    mysql = mysql & "t.vehiculo," & Chr$(10)
    mysql = mysql & "f.transporte " & Chr$(10)
    'mysql = mysql & "From factura f," & Chr$(10)
    mysql = mysql & "from " & cgusuario & " f," & Chr$(10)
    mysql = mysql & "transpor t" & Chr$(10)
    mysql = mysql & "where f.local='" & Trim("" & my_local) & "' " & Chr$(10)
    'If my_acu = "R" Then
    '   mysql = mysql & "and f.tipo='" & Trim("" & my_tipo1) & "' " & Chr$(10)
    'Else
    mysql = mysql & "and f.tipo='" & Trim("" & my_tipo) & "' " & Chr$(10)
    'End If
    mysql = mysql & "and f.serie='" & Trim("" & my_serie) & "' " & Chr$(10)
    mysql = mysql & "and f.numero='" & Trim("" & my_numero) & "'" & Chr$(10)
    mysql = mysql & "and F.tipo= f.TIPO" & Chr$(10)
    mysql = mysql & "and F.serie = f.SERIE" & Chr$(10)
    mysql = mysql & "and F.NUMERO = f.NUMERO" & Chr$(10)
    mysql = mysql & "and f.transporte = t.codigo" & Chr$(10)
 
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
                ReDim Preserve my_struc_Etransporte(UBound(my_struc_Etransporte) + 1)

            End If
      
            If mytablex.Fields("partida") <> "" Then
                my_struc_Etransporte(UBound(my_struc_Etransporte)).partida = mytablex.Fields("partida")
            Else
                my_struc_Etransporte(UBound(my_struc_Etransporte)).partida = ""

            End If
       
            If mytablex.Fields("destino") <> "" Then
                my_struc_Etransporte(UBound(my_struc_Etransporte)).destino = mytablex.Fields("destino")
            Else
                my_struc_Etransporte(UBound(my_struc_Etransporte)).destino = ""

            End If
    
            If mytablex.Fields("ruc") <> "" Then
                my_struc_Etransporte(UBound(my_struc_Etransporte)).RUC = mytablex.Fields("ruc")
            Else
                my_struc_Etransporte(UBound(my_struc_Etransporte)).RUC = ""

            End If
    
            If mytablex.Fields("nombreT") <> "" Then
                my_struc_Etransporte(UBound(my_struc_Etransporte)).nombreT = mytablex.Fields("nombreT")
            Else
                my_struc_Etransporte(UBound(my_struc_Etransporte)).nombreT = ""

            End If
   
            If mytablex.Fields("nombreC") <> "" Then
                my_struc_Etransporte(UBound(my_struc_Etransporte)).nombrec = mytablex.Fields("nombreC")
            Else
                my_struc_Etransporte(UBound(my_struc_Etransporte)).nombrec = ""

            End If
   
            If mytablex.Fields("fecha") <> "" Then
                my_struc_Etransporte(UBound(my_struc_Etransporte)).fecha = mytablex.Fields("fecha")
            Else
                my_struc_Etransporte(UBound(my_struc_Etransporte)).fecha = ""

            End If

            'aqui nuevo
            If mytablex.Fields("placa") <> "" Then
                my_struc_Etransporte(UBound(my_struc_Etransporte)).placa = mytablex.Fields("placa")
            Else
                my_struc_Etransporte(UBound(my_struc_Etransporte)).placa = ""

            End If
   
            If mytablex.Fields("licencia") <> "" Then
                my_struc_Etransporte(UBound(my_struc_Etransporte)).licencia = mytablex.Fields("licencia")
            Else
                my_struc_Etransporte(UBound(my_struc_Etransporte)).licencia = ""

            End If
  
            If mytablex.Fields("marca") <> "" Then
                my_struc_Etransporte(UBound(my_struc_Etransporte)).marca = mytablex.Fields("marca")
            Else
                my_struc_Etransporte(UBound(my_struc_Etransporte)).marca = ""

            End If
 
            If mytablex.Fields("vehiculo") <> "" Then
                my_struc_Etransporte(UBound(my_struc_Etransporte)).vehiculo = mytablex.Fields("vehiculo")
            Else
                my_struc_Etransporte(UBound(my_struc_Etransporte)).vehiculo = ""

            End If

            k = k + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
    Exit Function

End Function

Public Sub carga_transporte(my_struc_Etransporte() As struc_Etransporte, k As Integer)
           
    v = 13
    h = 0

    For I = 0 To k - 1
  
        objWorksheet.Cells(v, h + 2) = my_struc_Etransporte(I).nombrec & "-" & my_struc_Etransporte(I).nombreT 'enviar por
        objWorksheet.Cells(v, h + 6) = my_struc_Etransporte(I).RUC 'ruc
    
        objWorksheet.Cells(v + 1, h + 2) = my_struc_Etransporte(I).partida 'lugar entrega
  
        objWorksheet.Cells(v + 1, h + 6) = my_struc_Etransporte(I).fecha 'fecha entrega
  
        objWorksheet.Cells(v + 2, h + 2) = my_struc_Etransporte(I).destino
  
        v = v + 1
 
    Next I

    Exit Sub

End Sub

''*
Public Sub carga_impresion_sele(my_struc_cotizacion_total_excel() As struc_cotizacion_total_excel, _
                                D As Integer, _
                                my_moneda As String, _
                                my_transporte As String)
           
    v = 17
    h = 0
    sdx1 = 0
    sdx = 0
    subtotal = 0
    igv = 0

    On Error GoTo carga_impresion_sele

    For I = 0 To D - 1
        sdx = sdx + Val(my_struc_cotizacion_total_excel(I).cantidad)
        sdx1 = sdx1 + Val(my_struc_cotizacion_total_excel(I).total)
        subtotal = subtotal + Val(my_struc_cotizacion_total_excel(I).subtotal)
        igv = igv + Val(my_struc_cotizacion_total_excel(I).igv)
  
        objWorksheet.Cells(v, h + 1) = my_struc_cotizacion_total_excel(I).producto
        objWorksheet.Cells(v, h + 2) = "'" & my_struc_cotizacion_total_excel(I).descripcion
        objWorksheet.Cells(v, h + 3) = my_struc_cotizacion_total_excel(I).unidad
        objWorksheet.Cells(v, h + 4) = my_struc_cotizacion_total_excel(I).factor
        objWorksheet.Cells(v, h + 5) = my_struc_cotizacion_total_excel(I).cantidad
        objWorksheet.Cells(v, h + 6) = my_struc_cotizacion_total_excel(I).precio
        objWorksheet.Cells(v, h + 7) = my_struc_cotizacion_total_excel(I).total
  
        v = v + 1
    Next I

    'vemos en donde impacta 16/08/2017 pll
    If Mid(my_moneda, 1, 1) = "S" Then
        '  If my_moneda = "S" Then
        'inicio 16/08/2017 pll subtotal
        objWorksheet.Cells(v, h + 4).Font.bold = True
        objWorksheet.Cells(v, h + 4).Interior.color = RGB(215, 215, 0)  'resaltador
        objWorksheet.Cells(v, h + 4) = "Sub-TOTAL"
        
        objWorksheet.Cells(v, h + 7).Interior.color = RGB(215, 215, 0)  'resaltador
        objWorksheet.Cells(v, h + 7) = "" & Format(subtotal, "0.00") 'total subtotal
        'fin 16/08/2017 pll subtotal

        'inicio 17/08/2017 pll igv
        objWorksheet.Cells(v + 1, h + 4).Font.bold = True
        objWorksheet.Cells(v + 1, h + 4).Interior.color = RGB(215, 215, 0) 'resaltador
        objWorksheet.Cells(v + 1, h + 4) = "IGV"
        
        objWorksheet.Cells(v + 1, h + 7).Interior.color = RGB(215, 215, 0) 'resaltador
        objWorksheet.Cells(v + 1, h + 7) = "" & Format(igv, "0.00") 'total subtotal
        'fin 17/08/2017 pll igv

        objWorksheet.Cells(v + 2, h + 4).Font.bold = True
        objWorksheet.Cells(v + 2, h + 4).Interior.color = RGB(215, 215, 0) 'resaltador
        objWorksheet.Cells(v + 2, h + 4) = "TOTAL SOLES"
        
        objWorksheet.Cells(v + 2, h + 5).Interior.color = RGB(215, 215, 0) 'resaltador
        objWorksheet.Cells(v + 2, h + 5) = "" & Format(sdx, "0.00") 'total cantidad
   
        objWorksheet.Cells(v + 2, h + 7).Interior.color = RGB(215, 215, 0) 'resaltador
        objWorksheet.Cells(v + 2, h + 7) = "" & Format(sdx1, "0.00") 'total total
      
    Else
        'inicio 16/08/2017 pll subtotal
        objWorksheet.Cells(v, h + 4).Font.bold = True
        objWorksheet.Cells(v, h + 4).Interior.color = RGB(215, 215, 0)  'resaltador
        objWorksheet.Cells(v, h + 4) = "Sub-TOTAL"
        
        objWorksheet.Cells(v, h + 7).Interior.color = RGB(215, 215, 0)  'resaltador
        objWorksheet.Cells(v, h + 7) = "" & Format(subtotal, "0.00") 'total subtotal
        'fin 16/08/2017 pll subtotal

        'inicio 17/08/2017 pll igv
        objWorksheet.Cells(v + 1, h + 4).Font.bold = True
        objWorksheet.Cells(v + 1, h + 4).Interior.color = RGB(215, 215, 0) 'resaltador
        objWorksheet.Cells(v + 1, h + 4) = "IGV"
        
        objWorksheet.Cells(v + 1, h + 7).Interior.color = RGB(215, 215, 0) 'resaltador
        objWorksheet.Cells(v + 1, h + 7) = "" & Format(igv, "0.00") 'total subtotal
        'fin 17/08/2017 pll igv

        objWorksheet.Cells(v + 2, h + 4).Font.bold = True
        objWorksheet.Cells(v + 2, h + 4).Interior.color = RGB(215, 215, 0) 'resaltador
        objWorksheet.Cells(v + 2, h + 4) = "TOTAL DOLARES"
        
        objWorksheet.Cells(v + 2, h + 5).Interior.color = RGB(215, 215, 0) 'resaltador
        objWorksheet.Cells(v + 2, h + 5) = "" & Format(sdx, "0.00") 'total cantidad
   
        objWorksheet.Cells(v + 2, h + 7).Interior.color = RGB(215, 215, 0) 'resaltador
        objWorksheet.Cells(v + 2, h + 7) = "" & Format(sdx1, "0.00") 'total total

    End If
 
    'inicio 16/08/2017 pie de pagina
    objWorksheet.Cells(v + 3, h + 1).Font.bold = True
    objWorksheet.Cells(v + 3, h + 1).Select
    Selection.ColumnWidth = 14
             
    objWorksheet.Cells(v + 3, h + 1) = "Observaciones:"
    objWorksheet.Cells(v + 3, h + 2) = "______________________________________________"
    objWorksheet.Cells(v + 4, h + 2) = "______________________________________________"
 
    objWorksheet.Cells(v + 5, h + 1).Font.bold = True
    objWorksheet.Cells(v + 5, h + 1) = "Importante:"
    '
    objWorksheet.Cells(v + 6, h + 2) = "1. Colocar el número de orden de compra en las guías de remisión"
    objWorksheet.Cells(v + 7, h + 2) = "2. Facturar únicamente lo solicitado en esta orden de compra"
 
    objWorksheet.Cells(v + 8, h + 2) = "________________________"
    objWorksheet.Cells(v + 9, h + 2) = "       VoBo Logística"
 
    objWorksheet.Cells(v + 8, h + 4) = "________________________"
    objWorksheet.Cells(v + 9, h + 4) = "        VoBo Gerencia"
 
    'fin 16/08/2017 pie de pagina
  
    Exit Sub
 
carga_impresion_sele:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

End Sub

'Inicio 22/08/2017 pll
Public Sub cerra_excelR(my_file As String)

    ' objWorkBook.Protect Structure:=False, Windows:=False
    ' ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    If Mid(my_file, 1, 8) = "VGrafico" Then
        Worksheets(1).Protect UserInterfaceOnly:=True

    End If

    If Dir("C:\Reporte") = "" Then
        Call crea_directorio_excel

    End If

    new_file = "C:\Reporte\" & my_file & ".xls"
    objWorkBook.SaveAs "C:\Reporte\" & my_file & ".xls"
 
    Set objExcel = New Excel.Application
    Set xlibro = objExcel.Workbooks.Open("C:\Reporte\" & my_file & ".xls", vbMaximizedFocus)
    objExcel.Visible = True

    ' objWorkBook.Close

    Set objWorkBook = Nothing
    Set objWorksheet = Nothing
    oXL.Quit
    Set oXL = Nothing

End Sub

'Inicio 22/08/2017 pll
'inicio 23/08/2017 pll
Public Sub crea_directorio_excel()

    On Error GoTo crea_error_excel

    MkDir ("C:\Reporte")

crea_error_excel:
    Exit Sub

End Sub

'fin 23/08/2017 pll
Public Sub Carga_SActual(my_struc_saldo_actual() As struc_saldo_actual, k As Integer)

    v = 11
    h = 0
    sdx = 0
    sdx1 = 0
    sdx2 = 0

    For I = 0 To k - 1
        objWorksheet.Cells(v, h + 1) = my_struc_saldo_actual(I).familia
        objWorksheet.Cells(v, h + 2) = my_struc_saldo_actual(I).subfamilia
        objWorksheet.Cells(v, h + 3) = my_struc_saldo_actual(I).categoria
        objWorksheet.Cells(v, h + 4) = my_struc_saldo_actual(I).producto
        objWorksheet.Cells(v, h + 5) = my_struc_saldo_actual(I).descripcion
        objWorksheet.Cells(v, h + 6) = my_struc_saldo_actual(I).unidad
        objWorksheet.Cells(v, h + 7) = my_struc_saldo_actual(I).factor
        objWorksheet.Cells(v, h + 8) = my_struc_saldo_actual(I).cantidad
        my_cantidad = my_cantidad + my_struc_saldo_actual(I).cantidad
        objWorksheet.Cells(v, h + 9) = my_struc_saldo_actual(I).costou
        my_costou = my_costou + my_struc_saldo_actual(I).costou
        objWorksheet.Cells(v, h + 10) = my_struc_saldo_actual(I).total
        my_total = my_total + my_struc_saldo_actual(I).total
        objWorksheet.Cells(v, h + 11) = my_struc_saldo_actual(I).minimo
        v = v + 1
    Next I

    '**
    'aqui los sub-totales
    objWorksheet.Cells(v, h + 7).Font.bold = True
    objWorksheet.Cells(v, h + 7).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 7) = "TOTALES"
   
    objWorksheet.Cells(v, h + 8).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 8) = "" & Format(my_cantidad, "0.00") 'cantidad
   
    objWorksheet.Cells(v, h + 9).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 9) = "" & Format(my_costou, "0.00") 'precio unitario
   
    objWorksheet.Cells(v, h + 10).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 10) = "" & Format(my_total, "0.00") ''vtaxproc.total
    
End Sub

'inicio 13/12/2017 pll
Public Sub carga_Rventas(my_struc_Rventas() As struc_Rventas, k As Integer)

    v = 11
    h = 0

    Dim my_ndv_subtotal As Currency

    Dim my_ndv_total    As Currency

    Dim my_ndv_impuesto As Currency

    Dim my_ncv_subtotal As Currency

    Dim my_ncv_total    As Currency

    Dim my_ncv_impuesto As Currency

    Dim I               As Integer

    sdx = 0
    sdx1 = 0
    sdx2 = 0
    suma1 = 0
    suma2 = 0
    suma6 = 0
    ssuma6 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    ssuma4 = 0
    ssuma5 = 0
    suma7 = 0
    ssuma7 = 0
    suma8 = 0
    ssuma8 = 0
    suma9 = 0
    ssuma9 = 0

    On Error GoTo carga_Rventas

    Set objWorksheet = objWorkBook.Worksheets(1)

    'aqui el correlativo
    For I = 0 To k - 1
        objWorksheet.Cells(v, h + 1).Font.Size = 7 'aqui tamaño letra pll
        objWorksheet.Cells(v, h + 1) = "" & my_struc_Rventas(I).num_correlativo
 
        objWorksheet.Cells(v, h + 2).Font.Size = 7 'aqui tamaño letra pll
        objWorksheet.Cells(v, h + 2) = my_struc_Rventas(I).fecha
    
        objWorksheet.Cells(v, h + 3).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 3) = my_struc_Rventas(I).local
 
        If my_struc_Rventas(I).tipo = "1" Then
            objWorksheet.Cells(v, h + 4) = "TICK BOLETA"
            objWorksheet.Cells(v, h + 4).Font.Size = 7 'aqui tamaño letra
        ElseIf my_struc_Rventas(I).tipo = "2" Then
            objWorksheet.Cells(v, h + 4) = "TICK FACTURA"
            objWorksheet.Cells(v, h + 4).Font.Size = 7 'aqui tamaño letra
        Else
            objWorksheet.Cells(v, h + 4) = my_struc_Rventas(I).tipo
            objWorksheet.Cells(v, h + 4).Font.Size = 7 'aqui tamaño letra

        End If
 
        objWorksheet.Cells(v, h + 5).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 5) = my_struc_Rventas(I).serie
 
        objWorksheet.Cells(v, h + 6).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 6) = my_struc_Rventas(I).Numero
    
        objWorksheet.Cells(v, h + 7).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 7) = my_struc_Rventas(I).codigo
    
        objWorksheet.Cells(v, h + 8).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 8) = my_struc_Rventas(I).nombre

        If my_struc_Rventas(I).estado = "1" Then
            objWorksheet.Cells(v, h + 9).Font.Size = 4 'aqui tamaño letra
            objWorksheet.Cells(v, h + 9) = "ANULADO"
            objWorksheet.Range(Cells(v, h + 9), Cells(v, 17)).Interior.color = RGB(255, 164, 32)  'Color anaranjado
        Else
            objWorksheet.Cells(v, h + 9).Font.Size = 4 'aqui tamaño letra
            objWorksheet.Cells(v, h + 9) = "ACTIVADO"

        End If
 
        If my_struc_Rventas(I).tipo = "2" And my_struc_Rventas(I).acu = "D" And my_struc_Rventas(I).tipo1 = "NDV" Then 'venta normal que se ha convertido a nota Debito
       
            objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 10) = "+" & Format(my_struc_Rventas(I).subtotal - my_struc_Rventas(I).gravado, "0.00")

            'aqui es para subtotal
            my_ndv_subtotal = my_ndv_subtotal + Format(my_struc_Rventas(I).subtotal - my_struc_Rventas(I).gravado, "0.00")
                          
            'aqui es para el total
            my_ndv_total = my_ndv_total + Format(my_struc_Rventas(I).total, "0.00")
  
            'aqui el impuesto'
            my_ndv_impuesto = my_ndv_impuesto + Format(my_struc_Rventas(I).impuesto, "0.00")
  
            'aqui seria la diferencia my_ndv
            objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 11) = "+" & Format(my_struc_Rventas(I).gravado * Val(my_struc_Rventas(I).paridad), "0.00")
       
            objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 12) = "+" & Format(my_struc_Rventas(I).tisc * Val(my_struc_Rventas(I).paridad), "0.00")
       
            objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 13) = "+" & Format(my_struc_Rventas(I).impuesto, "0.00")
   
            objWorksheet.Cells(v, h + 14).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 14) = Format(my_struc_Rventas(I).total, "0.00")
   
            objWorksheet.Cells(v, h + 15).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 15) = "+" & Format(my_struc_Rventas(I).tivap * Format(my_struc_Rventas(I).paridad, "0.00"))
   
            objWorksheet.Cells(v, h + 16).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 16) = "+" & Format(my_struc_Rventas(I).percepcion * Format(my_struc_Rventas(I).paridad, "0.00"))
       
            objWorksheet.Cells(v, h + 17).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 17) = "+" & Format(my_struc_Rventas(I).servicioco * Format(my_struc_Rventas(I).paridad, "0.00"))
                                         
            objWorksheet.Cells(v, h + 18).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 18) = "+" & Format(my_struc_Rventas(I).tdetra * Format(my_struc_Rventas(I).paridad, "0.00"))
                                         
            '**FIN 07/12/2017 PLL
        ElseIf my_struc_Rventas(I).tipo = "NDV" And my_struc_Rventas(I).acu = "F" And my_struc_Rventas(I).tipo1 = "2" Then 'Alli abajo que se venta a nota de DEBITO Venta
       
            objWorksheet.Range(Cells(v, h + 10), Cells(v, 17)).Interior.color = RGB(127, 137, 0)  'Color verde
   
            objWorksheet.Cells(v, h + 10) = "'+" & Format(my_struc_Rventas(I).subtotal - my_struc_Rventas(I).gravado, "0.00")
       
            objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 11) = "'+" & Format(my_struc_Rventas(I).gravado * my_struc_Rventas(I).paridad, "0.00")
       
            objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 12) = "'+" & Format(my_struc_Rventas(I).tisc * my_struc_Rventas(I).paridad, "0.00")
            objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letra
   
            objWorksheet.Cells(v, h + 13) = "'+" & Format(my_struc_Rventas(I).impuesto, "0.00")
            objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra
       
            objWorksheet.Cells(v, h + 14) = "'+" & Format(my_struc_Rventas(I).total, "0.00")
            objWorksheet.Cells(v, h + 14).Font.Size = 7 'aqui tamaño letra
       
            objWorksheet.Cells(v, h + 15) = "'+" & Format(my_struc_Rventas(I).tivap * Format(my_struc_Rventas(I).paridad), "0.00")
            objWorksheet.Cells(v, h + 15).Font.Size = 7 'aqui tamaño letra
       
            objWorksheet.Cells(v, h + 16) = "'+" & Format(my_struc_Rventas(I).percepcion * Format(my_struc_Rventas(I).paridad), "0.00")
            objWorksheet.Cells(v, h + 16).Font.Size = 7 'aqui tamaño letra
       
            objWorksheet.Cells(v, h + 17) = "'+" & Format(my_struc_Rventas(I).servicioco * Format(my_struc_Rventas(I).paridad), "0.00")
            objWorksheet.Cells(v, h + 17).Font.Size = 7 'aqui tamaño letra
       
            objWorksheet.Cells(v, h + 18) = "'+" & Format(my_struc_Rventas(I).tdetra * Format(my_struc_Rventas(I).paridad), "0.00")
            objWorksheet.Cells(v, h + 18).Font.Size = 7 'aqui tamaño letra
   
            'nuevo aqui es para el subtotal
            objWorksheet.Cells(v, h + 19).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 19) = "'+" & Format(my_struc_Rventas(I).subtotal - my_ndv_subtotal, "0.00")
            my_ndv_subtotal = Format(my_struc_Rventas(I).subtotal - my_ndv_subtotal, "0.00")
        
            'aqui es para el total
            objWorksheet.Cells(v, h + 20).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 20) = "'+" & Format(my_struc_Rventas(I).total - my_ndv_total, "0.00")
            my_ndv_total = Format(my_struc_Rventas(I).total - my_ndv_total)
        
            'aqui es para el impuesto
            objWorksheet.Cells(v, h + 21).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 21) = "'+" & Format(my_struc_Rventas(I).impuesto - my_ndv_impuesto, "0.00")
            my_ndv_impuesto = Format(my_struc_Rventas(I).impuesto - my_ndv_impuesto)
        
            objWorksheet.Cells(v, h + 22).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 22) = "'" & my_struc_Rventas(I).fechasunat
        
            objWorksheet.Cells(v, h + 23).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 23) = "" & my_struc_Rventas(I).tipo1
        
            objWorksheet.Cells(v, h + 24).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 24) = "" & my_struc_Rventas(I).serie1
        
            objWorksheet.Cells(v, h + 25).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 25) = "" & my_struc_Rventas(I).numero1
        
            objWorksheet.Cells(v, h + 26).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 26) = "" & my_struc_Rventas(I).observa
            '**FIN 07/12/2017 PLL
        ElseIf my_struc_Rventas(I).tipo = "2" And my_struc_Rventas(I).acu = "D" And my_struc_Rventas(I).tipo1 = "0" Then 'Alli abajo 'aqui es venta normal
      
            objWorksheet.Cells(v, h + 10) = "'" & Format(Round(Val("" & my_struc_Rventas(I).subtotal - my_struc_Rventas(I).gravado), 2), "0.00")
            objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
      
            objWorksheet.Cells(v, h + 11) = "'" & Format(Round(Val("" & my_struc_Rventas(I).gravado * my_struc_Rventas(I).paridad), 2), "0.00")
                                     
            objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
      
            objWorksheet.Cells(v, h + 12) = "'" & Format(Round(Val("" & my_struc_Rventas(I).tisc * my_struc_Rventas(I).paridad), 2), "0.00")
            objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letra
      
            objWorksheet.Cells(v, h + 13) = "'" & Format(Round(Val("" & my_struc_Rventas(I).impuesto), 2), "0.00")
            objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra
      
            objWorksheet.Cells(v, h + 14) = "'" & Format(Round(Val("" & my_struc_Rventas(I).total), 2), "0.00")
            objWorksheet.Cells(v, h + 14).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 15) = "'" & Format(Round(Val("" & my_struc_Rventas(I).tivap * my_struc_Rventas(I).paridad), 2), "0.00")
            objWorksheet.Cells(v, h + 15).Font.Size = 7 'aqui tamaño letra
      
            objWorksheet.Cells(v, h + 16) = "'" & Format(Round(Val("" & my_struc_Rventas(I).percepcion * my_struc_Rventas(I).paridad), 2), "0.00")
            objWorksheet.Cells(v, h + 16).Font.Size = 7 'aqui tamaño letra
      
            objWorksheet.Cells(v, h + 17) = "'" & Format(Round(Val("" & my_struc_Rventas(I).servicioco * my_struc_Rventas(I).paridad), 2), "0.00")
            objWorksheet.Cells(v, h + 17).Font.Size = 7 'aqui tamaño letra
      
            objWorksheet.Cells(v, h + 18) = "'" & Format(Round(Val("" & my_struc_Rventas(I).tdetra * my_struc_Rventas(I).paridad), 2), "0.00")
            objWorksheet.Cells(v, h + 18).Font.Size = 7 'aqui tamaño letra
            '***inicio 10/11/2017 pll PARA NOTA CREDITO VENTAS
        ElseIf my_struc_Rventas(I).tipo = "2" And my_struc_Rventas(I).acu = "D" And my_struc_Rventas(I).tipo1 = "NCV" Then 'aqui es de ventas normal que se ha convertido a nota Credito
                      
            objWorksheet.Range(Cells(v, h + 10), Cells(v, 17)).Interior.color = RGB(255, 164, 32)  'Color anaranjado
   
            objWorksheet.Cells(v, h + 10) = "-" & Format(my_struc_Rventas(I).subtotal - my_struc_Rventas(I).gravado, "0.00")
            objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
        
            objWorksheet.Cells(v, h + 11) = "-" & Format(my_struc_Rventas(I).gravado * my_struc_Rventas(I).paridad, "0.00")
            objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
        
            objWorksheet.Cells(v, h + 12) = "-" & Format(my_struc_Rventas(I).tisc * my_struc_Rventas(I).paridad, "0.00")
            objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letraç
  
            objWorksheet.Cells(v, h + 13) = "-" & Format(my_struc_Rventas(I).impuesto, "0.00")
            objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra
     
            objWorksheet.Cells(v, h + 14) = "-" & Format(my_struc_Rventas(I).total, "0.00")
            objWorksheet.Cells(v, h + 14).Font.Size = 7 'aqui tamaño letra
        
            objWorksheet.Cells(v, h + 15) = "-" & Format(my_struc_Rventas(I).tivap * my_struc_Rventas(I).paridad, "0.00")
            objWorksheet.Cells(v, h + 15).Font.Size = 7 'aqui tamaño letra
      
            objWorksheet.Cells(v, h + 16) = "-" & Format(my_struc_Rventas(I).percepcion * my_struc_Rventas(I).paridad, "0.00")
            objWorksheet.Cells(v, h + 16).Font.Size = 7 'aqui tamaño letra
        
            objWorksheet.Cells(v, h + 17) = "-" & Format(my_struc_Rventas(I).servicioco * my_struc_Rventas(I).paridad, "0.00")
            objWorksheet.Cells(v, h + 17).Font.Size = 7 'aqui tamaño letra
      
            objWorksheet.Cells(v, h + 18) = "-" & Format(my_struc_Rventas(I).tdetra * my_struc_Rventas(I).paridad, "0.00")
            objWorksheet.Cells(v, h + 18).Font.Size = 7 'aqui tamaño letra
            '**FIN 07/12/2017 PLL
        ElseIf my_struc_Rventas(I).tipo = "NCV" And my_struc_Rventas(I).acu = "E" And my_struc_Rventas(I).tipo1 = "2" Then 'aqui documento ya convertido a nota Credito Ventas
     
            objWorksheet.Range(Cells(v, h + 10), Cells(v, 17)).Interior.color = RGB(255, 164, 32)  'Color anaranjado
          
            objWorksheet.Cells(v, h + 10) = "-" & Format(my_struc_Rventas(I).subtotal - my_struc_Rventas(I).gravado, "0.00")
            objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
        
            objWorksheet.Cells(v, h + 11) = "-" & Format(my_struc_Rventas(I).gravado * my_struc_Rventas(I).paridad, "0.00")
            objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
        
            objWorksheet.Cells(v, h + 12) = "-" & Format(my_struc_Rventas(I).tisc * my_struc_Rventas(I).paridad, "0.00")
            objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letra
     
            objWorksheet.Cells(v, h + 13) = "-" & Format(my_struc_Rventas(I).impuesto, "0.00")
            objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra
     
            objWorksheet.Cells(v, h + 14) = "-" & Format(my_struc_Rventas(I).total, "0.00")
            objWorksheet.Cells(v, h + 14).Font.Size = 7 'aqui tamaño letra
     
            objWorksheet.Cells(v, h + 15) = "-" & Format(my_struc_Rventas(I).tivap * my_struc_Rventas(I).paridad, "0.00")
            objWorksheet.Cells(v, h + 15).Font.Size = 7 'aqui tamaño letra
     
            objWorksheet.Cells(v, h + 16) = "-" & Format(my_struc_Rventas(I).percepcion * my_struc_Rventas(I).paridad, "0.00")
            objWorksheet.Cells(v, h + 16).Font.Size = 7 'aqui tamaño letra
     
            objWorksheet.Cells(v, h + 17) = "-" & Format(my_struc_Rventas(I).servicioco * my_struc_Rventas(I).paridad, "0.00")
            objWorksheet.Cells(v, h + 17).Font.Size = 7 'aqui tamaño letra
     
            objWorksheet.Cells(v, h + 18) = "-" & Format(my_struc_Rventas(I).tdetra * my_struc_Rventas(I).paridad, "0.00")
            objWorksheet.Cells(v, h + 18).Font.Size = 7 'aqui tamaño letra
 
            'nuevo aqui es para el subtotal
            objWorksheet.Cells(v, h + 19).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 19) = "'-" & Format(my_struc_Rventas(I).subtotal - my_ncv_subtotal, "0.00")
            my_ncv_subtotal = Format(my_struc_Rventas(I).subtotal - my_ncv_subtotal, "0.00")
        
            'aqui es para el total
            objWorksheet.Cells(v, h + 20).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 20) = "'-" & Format(my_struc_Rventas(I).total - my_ncv_total, "0.00")
            my_ncv_total = Format(my_struc_Rventas(I).total - my_ncv_total)
        
            'aqui es para el impuesto
            objWorksheet.Cells(v, h + 21).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 21) = "'-" & Format(my_struc_Rventas(I).impuesto - my_ncv_impuesto, "0.00")
            my_ncv_impuesto = Format(my_struc_Rventas(I).impuesto - my_ncv_impuesto)
        
            objWorksheet.Cells(v, h + 22).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 22) = "'" & my_struc_Rventas(I).fechasunat
        
            objWorksheet.Cells(v, h + 23).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 23) = "" & my_struc_Rventas(I).tipo1
        
            objWorksheet.Cells(v, h + 24).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 24) = "" & my_struc_Rventas(I).serie1
        
            objWorksheet.Cells(v, h + 25).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 25) = "" & my_struc_Rventas(I).numero1
        
            objWorksheet.Cells(v, h + 26).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 26) = "" & my_struc_Rventas(I).observa
            'inicio 15/12/2017 pll
        ElseIf my_struc_Rventas(I).tipo = "1" Then  'Alli abajo 'aqui es venta normal boleta
      
            objWorksheet.Cells(v, h + 10) = "'" & Format(Round(Val("" & my_struc_Rventas(I).subtotal - my_struc_Rventas(I).gravado), 2), "0.00")
            objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
      
            objWorksheet.Cells(v, h + 11) = "'" & Format(Round(Val("" & my_struc_Rventas(I).gravado * my_struc_Rventas(I).paridad), 2), "0.00")
            objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
      
            objWorksheet.Cells(v, h + 12) = "'" & Format(Round(Val("" & my_struc_Rventas(I).tisc * my_struc_Rventas(I).paridad), 2), "0.00")
            objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letra
      
            objWorksheet.Cells(v, h + 13) = "'" & Format(Round(Val("" & my_struc_Rventas(I).impuesto), 2), "0.00")
            objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra
      
            objWorksheet.Cells(v, h + 14) = "'" & Format(Round(Val("" & my_struc_Rventas(I).total), 2), "0.00")
            objWorksheet.Cells(v, h + 14).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 15) = "'" & Format(Round(Val("" & my_struc_Rventas(I).tivap * my_struc_Rventas(I).paridad), 2), "0.00")
            objWorksheet.Cells(v, h + 15).Font.Size = 7 'aqui tamaño letra
      
            objWorksheet.Cells(v, h + 16) = "'" & Format(Round(Val("" & my_struc_Rventas(I).percepcion * my_struc_Rventas(I).paridad), 2), "0.00")

            objWorksheet.Cells(v, h + 16).Font.Size = 7 'aqui tamaño letra
      
            objWorksheet.Cells(v, h + 17) = "'" & Format(Round(Val("" & my_struc_Rventas(I).servicioco * my_struc_Rventas(I).paridad), 2), "0.00")
            objWorksheet.Cells(v, h + 17).Font.Size = 7 'aqui tamaño letra
      
            objWorksheet.Cells(v, h + 18) = "'" & Format(Round(Val("" & my_struc_Rventas(I).tdetra * my_struc_Rventas(I).paridad), 2), "0.00")
            objWorksheet.Cells(v, h + 18).Font.Size = 7 'aqui tamaño letra

            'fin 15/12/2017 pll
        End If

        v = v + 1

        'aqui las acumulaciones ventas normal
        If my_struc_Rventas(I).tipo = "2" And my_struc_Rventas(I).acu = "D" And my_struc_Rventas(I).tipo1 = "NDV" Then 'aqui es de ventas normal convertido a nota Debito Venta
            'aqui es acumulacion de ventas normal
            'aqui es para sl subtotal
            ssuma1 = suma1 + Val("" & my_struc_Rventas(I).subtotal)
     
            'aqui es parael impuesto
            ssuma4 = suma4 + Val("" & my_struc_Rventas(I).impuesto)
     
            'aqui es para el total
            ssuma5 = suma5 + Val("" & my_struc_Rventas(I).total)
       
            'aqui es para el tivap
            ssuma6 = suma6 + Format(my_struc_Rventas(I).tivap * my_struc_Rventas(I).paridad, "0.00")
                 
            'aqui es para isc
            ssuma3 = suma3 + Val(my_struc_Rventas(I).tisc)
       
            'aqui es para percepcion
            ssuma7 = suma7 + Format(my_struc_Rventas(I).percepcion * my_struc_Rventas(I).paridad, "0.00")
                 
            'aqui es para el servicio
            ssuma8 = suma8 + Format(my_struc_Rventas(I).servicioco * my_struc_Rventas(I).paridad, "0.00")
                   
            'aqui es para detraccion
            ssuma9 = suma9 + Format(my_struc_Rventas(I).tdetra * my_struc_Rventas(I).paridad, "0.00")
                  
            'abajo no se donde lo conecta pll
            suma2 = suma2 + Format(my_struc_Rventas(I).gravado * my_struc_Rventas(I).paridad, "0.00")
            ssuma2 = ssuma2 + Format(my_struc_Rventas(I).gravado * my_struc_Rventas(I).paridad, "0.00")
                     
        End If

        'inicio 09/12/2017 PLL
        If my_struc_Rventas(I).tipo = "NDV" And my_struc_Rventas(I).acu = "F" And my_struc_Rventas(I).tipo1 = "2" Then 'aqui es de ventas normal convertido a nota Debito Venta
            'para el subtotal
            suma1 = suma1 + my_ndv_subtotal
            ssuma1 = ssuma1 + suma1
   
            'aqui es para el total
            suma5 = suma5 + my_ndv_total
            ssuma5 = ssuma5 + suma5
            'aqui es para el impuesto
            suma4 = suma4 + my_ndv_impuesto
            ssuma4 = ssuma4 + suma4

        End If

        'Fin 09/12/2017 PLL
        'Inicio 12/12/2017 pll
        If my_struc_Rventas(I).tipo = "2" And my_struc_Rventas(I).acu = "D" And my_struc_Rventas(I).tipo1 = "0" Then 'Alli abajo 'aqui es venta normal
    
            'aqui es acumulacion de ventas normal
            'aqui es para sl subtotal
            ssuma1 = ssuma1 + Format(my_struc_Rventas(I).subtotal)
            'aqui es parael impuesto
            ssuma4 = ssuma4 + Format(my_struc_Rventas(I).impuesto)
            'aqui es para el total
            ssuma5 = ssuma5 + Format(my_struc_Rventas(I).total)
  
            'aqui es para el tivap
            ssuma6 = suma6 + Format(my_struc_Rventas(I).tivap * my_struc_Rventas(I).paridad, "0.00")
                 
            'aqui es para isc
            ssuma3 = suma3 + Val(my_struc_Rventas(I).tisc)
       
            'aqui es para percepcion
            ssuma7 = suma7 + Format(my_struc_Rventas(I).percepcion * my_struc_Rventas(I).paridad, "0.00")
                 
            'aqui es para el servicio
            ssuma8 = suma8 + Format(my_struc_Rventas(I).servicioco * my_struc_Rventas(I).paridad, "0.00")
                   
            'aqui es para detraccion
            ssuma9 = suma9 + Format(my_struc_Rventas(I).tdetra * my_struc_Rventas(I).paridad, "0.00")

        End If

        'fin 12/12/2017 pll
        'inicio 15/12/2017 pll
        If my_struc_Rventas(I).tipo = "1" Then  'Alli abajo 'aqui es venta normal boleta
            'aqui es acumulacion de ventas normal
            'aqui es para sl subtotal
            ssuma1 = ssuma1 + Format(my_struc_Rventas(I).subtotal)
            'aqui es parael impuesto
            ssuma4 = ssuma4 + Format(my_struc_Rventas(I).impuesto)
            'aqui es para el total
            ssuma5 = ssuma5 + Format(my_struc_Rventas(I).total)

            'aqui es para el tivap
            ssuma6 = suma6 + Format(my_struc_Rventas(I).tivap * my_struc_Rventas(I).paridad, "0.00")
                 
            'aqui es para isc
            ssuma3 = suma3 + Val(my_struc_Rventas(I).tisc)
       
            'aqui es para percepcion
            ssuma7 = suma7 + Format(my_struc_Rventas(I).percepcion * my_struc_Rventas(I).paridad, "0.00")
                 
            'aqui es para el servicio
            ssuma8 = suma8 + Format(my_struc_Rventas(I).servicioco * my_struc_Rventas(I).paridad, "0.00")
                   
            'aqui es para detraccion
            ssuma9 = suma9 + Format(my_struc_Rventas(I).tdetra * my_struc_Rventas(I).paridad, "0.00")

        End If

        'fin 15/12/2017 pll
    Next I

    '*****************************
    'aqui los sub-totales
    h = 3
    v = v + 1
    'aqui los totales
    'provemos anteriormente tambien estaba asi
    objWorksheet.Cells(v, h + 6).Font.bold = True
    objWorksheet.Cells(v, h + 6).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 6) = "TOTALES"
   
    objWorksheet.Cells(v, h + 7).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 7) = "'" & Format(Round(Val(ssuma1), 2), "0.00")   'subtotal

    objWorksheet.Cells(v, h + 8).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 8) = "" & Format(ssuma2, "0.00") 'precio unitario
   
    objWorksheet.Cells(v, h + 9).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 9) = "" & Format(ssuma3, "0.00") 'isc
   
    objWorksheet.Cells(v, h + 10).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 10) = "'" & Format(ssuma4, "0.00") 'impuesto IGV
   
    objWorksheet.Cells(v, h + 11).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 11) = "'" & Format(Round(Val(ssuma5), 2), "0.00")  'costo total

    objWorksheet.Cells(v, h + 12).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 12) = "" & Format(ssuma6, "0.00") 'tivap
   
    objWorksheet.Cells(v, h + 13).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 13) = "" & Format(ssuma7, "0.00") 'percepcion

    objWorksheet.Cells(v, h + 14).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 14) = "" & Format(ssuma8, "0.00") 'servicio

    objWorksheet.Cells(v, h + 15).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 15) = "" & Format(ssuma9, "0.00") 'detraccion

    Exit Sub
 
carga_Rventas:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

    End Select

End Sub

Public Sub carga_Rcompras(my_struc_Rventas() As struc_Rventas, k As Integer)

    v = 11
    h = 0

    Dim my_ndv_subtotal As Currency

    Dim my_ndv_total    As Currency

    Dim my_ndv_impuesto As Currency

    Dim my_ncv_subtotal As Currency

    Dim my_ncv_total    As Currency

    Dim my_ncv_impuesto As Currency

    Dim I               As Integer

    sdx = 0
    sdx1 = 0
    sdx2 = 0
    suma1 = 0
    suma2 = 0
    suma6 = 0
    ssuma6 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    ssuma4 = 0
    ssuma5 = 0
    suma7 = 0
    ssuma7 = 0
    suma8 = 0
    ssuma8 = 0
    suma9 = 0
    ssuma9 = 0

    On Error GoTo carga_Rcompras

    Set objWorksheet = objWorkBook.Worksheets(1)

    'aqui el correlativo
    For I = 0 To k - 1
        objWorksheet.Cells(v, h + 1).Font.Size = 7 'aqui tamaño letra pll
        objWorksheet.Cells(v, h + 1) = "" & my_struc_Rventas(I).num_correlativo
 
        objWorksheet.Cells(v, h + 2).Font.Size = 7 'aqui tamaño letra pll
        objWorksheet.Cells(v, h + 2) = my_struc_Rventas(I).fecha
    
        objWorksheet.Cells(v, h + 3).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 3) = my_struc_Rventas(I).local
 
        If my_struc_Rventas(I).tipo = "FC" Then
            objWorksheet.Cells(v, h + 4) = "FACTURA DE COMPRA"
            objWorksheet.Cells(v, h + 4).Font.Size = 7 'aqui tamaño letra
        ElseIf my_struc_Rventas(I).tipo = "BC" Then
            objWorksheet.Cells(v, h + 4) = "BOLETA DE COMPRA"
            objWorksheet.Cells(v, h + 4).Font.Size = 7 'aqui tamaño letra
        Else
            objWorksheet.Cells(v, h + 4) = my_struc_Rventas(I).tipo
            objWorksheet.Cells(v, h + 4).Font.Size = 7 'aqui tamaño letra

        End If
 
        objWorksheet.Cells(v, h + 5).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 5) = my_struc_Rventas(I).serie
 
        objWorksheet.Cells(v, h + 6).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 6) = my_struc_Rventas(I).Numero
    
        objWorksheet.Cells(v, h + 7).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 7) = my_struc_Rventas(I).codigo
    
        objWorksheet.Cells(v, h + 8).Font.Size = 7 'aqui tamaño letra
        objWorksheet.Cells(v, h + 8) = my_struc_Rventas(I).nombre

        If my_struc_Rventas(I).estado = "1" Then
            objWorksheet.Cells(v, h + 9).Font.Size = 4 'aqui tamaño letra
            objWorksheet.Cells(v, h + 9) = "ANULADO"
            objWorksheet.Range(Cells(v, h + 9), Cells(v, 17)).Interior.color = RGB(255, 164, 32)  'Color anaranjado
        Else
            objWorksheet.Cells(v, h + 9).Font.Size = 4 'aqui tamaño letra
            objWorksheet.Cells(v, h + 9) = "ACTIVADO"

        End If
 
        'If my_struc_Rventas(i).tipo = "2" And my_struc_Rventas(i).acu = "D" And _
        '   my_struc_Rventas(i).tipo1 = "NDV" Then 'aqui es de ventas normal que se ha convertido a nota Debito
       
        '   objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
        '   objWorksheet.Cells(v, h + 10) = "+" & Format(my_struc_Rventas(i).subtotal - _
            my_struc_Rventas(i).gravado, "0.00")

        'aqui es para subtotal
        '   my_ndv_subtotal = my_ndv_subtotal + Format(my_struc_Rventas(i).subtotal - _
            my_struc_Rventas(i).gravado, "0.00")
                          
        'aqui es para el total
        '  my_ndv_total = my_ndv_total + Format(my_struc_Rventas(i).total, "0.00")
  
        'aqui el impuesto'
        '  my_ndv_impuesto = my_ndv_impuesto + Format(my_struc_Rventas(i).impuesto, "0.00")
  
        'aqui seria la diferencia my_ndv
        '  objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
        '  objWorksheet.Cells(v, h + 11) = "+" & Format(my_struc_Rventas(i).gravado * Val(my_struc_Rventas(i).paridad), "0.00")
       
        '  objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letra
        '  objWorksheet.Cells(v, h + 12) = "+" & Format(my_struc_Rventas(i).tisc * Val(my_struc_Rventas(i).paridad), "0.00")
       
        '  objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra
        '  objWorksheet.Cells(v, h + 13) = "+" & Format(my_struc_Rventas(i).impuesto, "0.00")
   
        '  objWorksheet.Cells(v, h + 14).Font.Size = 7 'aqui tamaño letra
        '  objWorksheet.Cells(v, h + 14) = Format(my_struc_Rventas(i).total, "0.00")
   
        '  objWorksheet.Cells(v, h + 15).Font.Size = 7 'aqui tamaño letra
        '  objWorksheet.Cells(v, h + 15) = "+" & Format(my_struc_Rventas(i).tivap * _
           Format(my_struc_Rventas(i).paridad, "0.00"))
   
        '  objWorksheet.Cells(v, h + 16).Font.Size = 7 'aqui tamaño letra
        '  objWorksheet.Cells(v, h + 16) = "+" & Format(my_struc_Rventas(i).percepcion * _
           Format(my_struc_Rventas(i).paridad, "0.00"))
       
        '  objWorksheet.Cells(v, h + 17).Font.Size = 7 'aqui tamaño letra
        '  objWorksheet.Cells(v, h + 17) = "+" & Format(my_struc_Rventas(i).servicioco * _
           Format(my_struc_Rventas(i).paridad, "0.00"))
                                         
        '  objWorksheet.Cells(v, h + 18).Font.Size = 7 'aqui tamaño letra
        '  objWorksheet.Cells(v, h + 18) = "+" & Format(my_struc_Rventas(i).tdetra * _
           Format(my_struc_Rventas(i).paridad, "0.00"))
                                         
        '**FIN 07/12/2017 PLL
        'ElseIf my_struc_Rventas(i).tipo = "NDV" And my_struc_Rventas(i).acu = "F" And _
        '       my_struc_Rventas(i).tipo1 = "2" Then 'Alli abajo que se venta a nota de DEBITO Venta
       
        '   objWorksheet.Range(Cells(v, h + 10), Cells(v, 17)).Interior.color = RGB(127, 137, 0)  'Color verde
   
        '   objWorksheet.Cells(v, h + 10) = "'+" & Format(my_struc_Rventas(i).subtotal - _
            my_struc_Rventas(i).gravado, "0.00")
       
        '   objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
        '   objWorksheet.Cells(v, h + 11) = "'+" & Format(my_struc_Rventas(i).gravado * _
            my_struc_Rventas(i).paridad, "0.00")
       
        '   objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
        '   objWorksheet.Cells(v, h + 12) = "'+" & Format(my_struc_Rventas(i).tisc * _
            my_struc_Rventas(i).paridad, "0.00")
        '   objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letra
   
        '   objWorksheet.Cells(v, h + 13) = "'+" & Format(my_struc_Rventas(i).impuesto, "0.00")
        '   objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra
       
        '   objWorksheet.Cells(v, h + 14) = "'+" & Format(my_struc_Rventas(i).total, "0.00")
        '   objWorksheet.Cells(v, h + 14).Font.Size = 7 'aqui tamaño letra
       
        '   objWorksheet.Cells(v, h + 15) = "'+" & Format(my_struc_Rventas(i).tivap * _
            Format(my_struc_Rventas(i).paridad), "0.00")
        '   objWorksheet.Cells(v, h + 15).Font.Size = 7 'aqui tamaño letra
       
        '   objWorksheet.Cells(v, h + 16) = "'+" & Format(my_struc_Rventas(i).percepcion * _
            Format(my_struc_Rventas(i).paridad), "0.00")
        '   objWorksheet.Cells(v, h + 16).Font.Size = 7 'aqui tamaño letra
       
        '   objWorksheet.Cells(v, h + 17) = "'+" & Format(my_struc_Rventas(i).servicioco * _
            Format(my_struc_Rventas(i).paridad), "0.00")
        '   objWorksheet.Cells(v, h + 17).Font.Size = 7 'aqui tamaño letra
       
        '   objWorksheet.Cells(v, h + 18) = "'+" & Format(my_struc_Rventas(i).tdetra * _
            Format(my_struc_Rventas(i).paridad), "0.00")
        '   objWorksheet.Cells(v, h + 18).Font.Size = 7 'aqui tamaño letra
   
        'nuevo aqui es para el subtotal
        '   objWorksheet.Cells(v, h + 19).Font.Size = 7 'aqui tamaño letra
        '   objWorksheet.Cells(v, h + 19) = "'+" & Format(my_struc_Rventas(i).subtotal - my_ndv_subtotal, "0.00")
        '   my_ndv_subtotal = Format(my_struc_Rventas(i).subtotal - my_ndv_subtotal, "0.00")
        
        'aqui es para el total
        '   objWorksheet.Cells(v, h + 20).Font.Size = 7 'aqui tamaño letra
        '   objWorksheet.Cells(v, h + 20) = "'+" & Format(my_struc_Rventas(i).total - my_ndv_total, "0.00")
        '   my_ndv_total = Format(my_struc_Rventas(i).total - my_ndv_total)
        
        'aqui es para el impuesto
        '   objWorksheet.Cells(v, h + 21).Font.Size = 7 'aqui tamaño letra
        '   objWorksheet.Cells(v, h + 21) = "'+" & Format(my_struc_Rventas(i).impuesto - my_ndv_impuesto, "0.00")
        '   my_ndv_impuesto = Format(my_struc_Rventas(i).impuesto - my_ndv_impuesto)
        
        '   objWorksheet.Cells(v, h + 22).Font.Size = 7 'aqui tamaño letra
        '   objWorksheet.Cells(v, h + 22) = "'" & my_struc_Rventas(i).fechasunat
        
        '   objWorksheet.Cells(v, h + 23).Font.Size = 7 'aqui tamaño letra
        '   objWorksheet.Cells(v, h + 23) = "" & my_struc_Rventas(i).tipo1
        
        '   objWorksheet.Cells(v, h + 24).Font.Size = 7 'aqui tamaño letra
        '   objWorksheet.Cells(v, h + 24) = "" & my_struc_Rventas(i).serie1
        
        '   objWorksheet.Cells(v, h + 25).Font.Size = 7 'aqui tamaño letra
        '   objWorksheet.Cells(v, h + 25) = "" & my_struc_Rventas(i).numero1
        
        '   objWorksheet.Cells(v, h + 26).Font.Size = 7 'aqui tamaño letra
        '   objWorksheet.Cells(v, h + 26) = "" & my_struc_Rventas(i).observa
        '**FIN 07/12/2017 PLL
        'ElseIf my_struc_Rventas(i).tipo = "FC" And my_struc_Rventas(i).acu = "K" And _
        '      my_struc_Rventas(i).tipo1 = "0" Then 'Alli abajo 'aqui es venta normal
      
        If my_struc_Rventas(I).tipo = "FC" And my_struc_Rventas(I).acu = "K" And my_struc_Rventas(I).tipo1 = "0" Then 'Alli abajo 'aqui es venta normal
      
            objWorksheet.Cells(v, h + 10) = Format(my_struc_Rventas(I).subtotal - my_struc_Rventas(I).gravado, "0.00")
            objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
      
            objWorksheet.Cells(v, h + 11) = Format(my_struc_Rventas(I).gravado * my_struc_Rventas(I).paridad, "0.00")
            objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
      
            objWorksheet.Cells(v, h + 12) = Format(my_struc_Rventas(I).tisc * my_struc_Rventas(I).paridad, "0.00")
            objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letra
      
            objWorksheet.Cells(v, h + 13) = Format(my_struc_Rventas(I).impuesto, "0.00")
            objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra
      
            objWorksheet.Cells(v, h + 14) = Format(my_struc_Rventas(I).total, "0.00")
            objWorksheet.Cells(v, h + 14).Font.Size = 7 'aqui tamaño letra
            objWorksheet.Cells(v, h + 15) = Format(my_struc_Rventas(I).tivap * my_struc_Rventas(I).paridad, "0.00")
            objWorksheet.Cells(v, h + 15).Font.Size = 7 'aqui tamaño letra
      
            objWorksheet.Cells(v, h + 16) = Format(my_struc_Rventas(I).percepcion * my_struc_Rventas(I).paridad, "0.00")
            objWorksheet.Cells(v, h + 16).Font.Size = 7 'aqui tamaño letra
      
            objWorksheet.Cells(v, h + 17) = Format(my_struc_Rventas(I).servicioco * my_struc_Rventas(I).paridad, "0.00")
            objWorksheet.Cells(v, h + 17).Font.Size = 7 'aqui tamaño letra
      
            objWorksheet.Cells(v, h + 18) = Format(my_struc_Rventas(I).tdetra * my_struc_Rventas(I).paridad, "0.00")
            objWorksheet.Cells(v, h + 18).Font.Size = 7 'aqui tamaño letra
            '***inicio 10/11/2017 pll PARA NOTA CREDITO VENTAS
            'ElseIf my_struc_Rventas(i).tipo = "2" And my_struc_Rventas(i).acu = "D" And _
             my_struc_Rventas(i).tipo1 = "NCV" Then 'aqui es de ventas normal que se ha convertido a nota Credito
                      
            '  objWorksheet.Range(Cells(v, h + 10), Cells(v, 17)).Interior.color = RGB(255, 164, 32)  'Color anaranjado
   
            '  objWorksheet.Cells(v, h + 10) = "-" & Format(my_struc_Rventas(i).subtotal - _
               my_struc_Rventas(i).gravado, "0.00")
            '  objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
        
            '  objWorksheet.Cells(v, h + 11) = "-" & Format(my_struc_Rventas(i).gravado * _
               my_struc_Rventas(i).paridad, "0.00")
            '  objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
        
            '  objWorksheet.Cells(v, h + 12) = "-" & Format(my_struc_Rventas(i).tisc * _
               my_struc_Rventas(i).paridad, "0.00")
            '  objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letraç
  
            '  objWorksheet.Cells(v, h + 13) = "-" & Format(my_struc_Rventas(i).impuesto, "0.00")
            '  objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra
     
            '  objWorksheet.Cells(v, h + 14) = "-" & Format(my_struc_Rventas(i).total, "0.00")
            '  objWorksheet.Cells(v, h + 14).Font.Size = 7 'aqui tamaño letra
        
            '  objWorksheet.Cells(v, h + 15) = "-" & Format(my_struc_Rventas(i).tivap * _
               my_struc_Rventas(i).paridad, "0.00")
            '  objWorksheet.Cells(v, h + 15).Font.Size = 7 'aqui tamaño letra
      
            '  objWorksheet.Cells(v, h + 16) = "-" & Format(my_struc_Rventas(i).percepcion * _
               my_struc_Rventas(i).paridad, "0.00")
            '  objWorksheet.Cells(v, h + 16).Font.Size = 7 'aqui tamaño letra
        
            '  objWorksheet.Cells(v, h + 17) = "-" & Format(my_struc_Rventas(i).servicioco * _
               my_struc_Rventas(i).paridad, "0.00")
            ' objWorksheet.Cells(v, h + 17).Font.Size = 7 'aqui tamaño letra
      
            ' objWorksheet.Cells(v, h + 18) = "-" & Format(my_struc_Rventas(i).tdetra * _
              my_struc_Rventas(i).paridad, "0.00")
            '  objWorksheet.Cells(v, h + 18).Font.Size = 7 'aqui tamaño letra
            '**FIN 07/12/2017 PLL
            'ElseIf my_struc_Rventas(i).tipo = "NCV" And my_struc_Rventas(i).acu = "E" And _
             my_struc_Rventas(i).tipo1 = "2" Then 'aqui documento ya convertido a nota Credito Ventas
     
            '   objWorksheet.Range(Cells(v, h + 10), Cells(v, 17)).Interior.color = RGB(255, 164, 32)  'Color anaranjado
          
            '   objWorksheet.Cells(v, h + 10) = "-" & Format(my_struc_Rventas(i).subtotal - _
                my_struc_Rventas(i).gravado, "0.00")
            '   objWorksheet.Cells(v, h + 10).Font.Size = 7 'aqui tamaño letra
        
            '   objWorksheet.Cells(v, h + 11) = "-" & Format(my_struc_Rventas(i).gravado * _
                my_struc_Rventas(i).paridad, "0.00")
            '  objWorksheet.Cells(v, h + 11).Font.Size = 7 'aqui tamaño letra
        
            '  objWorksheet.Cells(v, h + 12) = "-" & Format(my_struc_Rventas(i).tisc * _
               my_struc_Rventas(i).paridad, "0.00")
            '  objWorksheet.Cells(v, h + 12).Font.Size = 7 'aqui tamaño letra
     
            '  objWorksheet.Cells(v, h + 13) = "-" & Format(my_struc_Rventas(i).impuesto, "0.00")
            '  objWorksheet.Cells(v, h + 13).Font.Size = 7 'aqui tamaño letra
     
            '  objWorksheet.Cells(v, h + 14) = "-" & Format(my_struc_Rventas(i).total, "0.00")
            '  objWorksheet.Cells(v, h + 14).Font.Size = 7 'aqui tamaño letra
     
            '  objWorksheet.Cells(v, h + 15) = "-" & Format(my_struc_Rventas(i).tivap * _
               my_struc_Rventas(i).paridad, "0.00")
            ' objWorksheet.Cells(v, h + 15).Font.Size = 7 'aqui tamaño letra
     
            '  objWorksheet.Cells(v, h + 16) = "-" & Format(my_struc_Rventas(i).percepcion * _
               my_struc_Rventas(i).paridad, "0.00")
            '  objWorksheet.Cells(v, h + 16).Font.Size = 7 'aqui tamaño letra
     
            '  objWorksheet.Cells(v, h + 17) = "-" & Format(my_struc_Rventas(i).servicioco * _
               my_struc_Rventas(i).paridad, "0.00")
            '  objWorksheet.Cells(v, h + 17).Font.Size = 7 'aqui tamaño letra
     
            '  objWorksheet.Cells(v, h + 18) = "-" & Format(my_struc_Rventas(i).tdetra * _
               my_struc_Rventas(i).paridad, "0.00")
            '  objWorksheet.Cells(v, h + 18).Font.Size = 7 'aqui tamaño letra
 
            'nuevo aqui es para el subtotal
            '  objWorksheet.Cells(v, h + 19).Font.Size = 7 'aqui tamaño letra
            '  objWorksheet.Cells(v, h + 19) = "'-" & Format(my_struc_Rventas(i).subtotal - my_ncv_subtotal, "0.00")
            '  my_ncv_subtotal = Format(my_struc_Rventas(i).subtotal - my_ncv_subtotal, "0.00")
        
            'aqui es para el total
            '  objWorksheet.Cells(v, h + 20).Font.Size = 7 'aqui tamaño letra
            '  objWorksheet.Cells(v, h + 20) = "'-" & Format(my_struc_Rventas(i).total - my_ncv_total, "0.00")
            '  my_ncv_total = Format(my_struc_Rventas(i).total - my_ncv_total)
        
            'aqui es para el impuesto
            '  objWorksheet.Cells(v, h + 21).Font.Size = 7 'aqui tamaño letra
            '  objWorksheet.Cells(v, h + 21) = "'-" & Format(my_struc_Rventas(i).impuesto - my_ncv_impuesto, "0.00")
            '  my_ncv_impuesto = Format(my_struc_Rventas(i).impuesto - my_ncv_impuesto)
        
            '  objWorksheet.Cells(v, h + 22).Font.Size = 7 'aqui tamaño letra
            '  objWorksheet.Cells(v, h + 22) = "'" & my_struc_Rventas(i).fechasunat
        
            '  objWorksheet.Cells(v, h + 23).Font.Size = 7 'aqui tamaño letra
            '  objWorksheet.Cells(v, h + 23) = "" & my_struc_Rventas(i).tipo1
        
            '  objWorksheet.Cells(v, h + 24).Font.Size = 7 'aqui tamaño letra
            '  objWorksheet.Cells(v, h + 24) = "" & my_struc_Rventas(i).serie1
        
            '  objWorksheet.Cells(v, h + 25).Font.Size = 7 'aqui tamaño letra
            '  objWorksheet.Cells(v, h + 25) = "" & my_struc_Rventas(i).numero1
        
            '  objWorksheet.Cells(v, h + 26).Font.Size = 7 'aqui tamaño letra
            '  objWorksheet.Cells(v, h + 26) = "" & my_struc_Rventas(i).observa
        End If

        v = v + 1
        'aqui las acumulaciones ventas normal
        'If my_struc_Rventas(i).tipo = "2" And my_struc_Rventas(i).acu = "D" And _
         my_struc_Rventas(i).tipo1 = "NDV" Then 'aqui es de ventas normal convertido a nota Debito Venta
        'aqui es acumulacion de ventas normal
        'aqui es para sl subtotal
        '  ssuma1 = suma1 + Val("" & my_struc_Rventas(i).subtotal)
     
        'aqui es parael impuesto
        '  ssuma4 = suma4 + Val("" & my_struc_Rventas(i).impuesto)
     
        'aqui es para el total
        '  ssuma5 = suma5 + Val("" & my_struc_Rventas(i).total)
       
        '  suma2 = suma2 + Format(my_struc_Rventas(i).gravado * _
           my_struc_Rventas(i).paridad, "0.00")
        '  ssuma2 = ssuma2 + Format(my_struc_Rventas(i).gravado * _
           my_struc_Rventas(i).paridad, "0.00")
    
        '  suma6 = suma6 + Format(my_struc_Rventas(i).tivap * _
           my_struc_Rventas(i).paridad, "0.00")
        '  ssuma6 = ssuma6 + Format(my_struc_Rventas(i).tivap * _
           my_struc_Rventas(i).paridad, "0.00")
                
        '  suma3 = suma3 + Format(my_struc_Rventas(i).tisc * _
           my_struc_Rventas(i).paridad, "0.00")
     
        '  ssuma3 = ssuma3 + Format(my_struc_Rventas(i).tisc * _
           my_struc_Rventas(i).paridad, "0.00")
     
        ' suma7 = suma7 + Format(my_struc_Rventas(i).percepcion * _
          my_struc_Rventas(i).paridad, "0.00")
        ' ssuma7 = ssuma7 + Format(my_struc_Rventas(i).percepcion * _
          my_struc_Rventas(i).paridad, "0.00")
                  
        ' suma8 = suma8 + Format(my_struc_Rventas(i).servicioco * _
          my_struc_Rventas(i).paridad, "0.00")
        ' ssuma8 = ssuma8 + Format(my_struc_Rventas(i).servicioco * _
          my_struc_Rventas(i).paridad, "0.00")
      
        'suma9 = suma9 + Format(my_struc_Rventas(i).tdetra * _
         my_struc_Rventas(i).paridad, "0.00")
        ' ssuma9 = ssuma9 + Format(my_struc_Rventas(i).tdetra * _
          my_struc_Rventas(i).paridad, "0.00")
                  
        'End If
        'inicio 09/12/2017 PLL
        'If my_struc_Rventas(i).tipo = "NDV" And my_struc_Rventas(i).acu = "F" And _
         my_struc_Rventas(i).tipo1 = "2" Then 'aqui es de ventas normal convertido a nota Debito Venta
        'para el subtotal
        '   suma1 = suma1 + my_ndv_subtotal
        '   ssuma1 = ssuma1 + suma1
   
        'aqui es para el total
        '   suma5 = suma5 + my_ndv_total
        '   ssuma5 = ssuma5 + suma5
        'aqui es para el impuesto
        '  suma4 = suma4 + my_ndv_impuesto
        '  ssuma4 = ssuma4 + suma4
        'End If
        'Fin 09/12/2017 PLL
        'Inicio 12/12/2017 pll
        If my_struc_Rventas(I).tipo = "FC" And my_struc_Rventas(I).acu = "K" And my_struc_Rventas(I).tipo1 = "0" Then 'Alli abajo 'aqui es venta normal
    
            'aqui es acumulacion de ventas normal
            'aqui es para sl subtotal
            'ssuma1 = suma1 + Val("" & mytablex.Fields("subtotal"))
            ssuma1 = ssuma1 + Format(my_struc_Rventas(I).subtotal)
     
            'aqui es parael impuesto
            ssuma4 = ssuma4 + Format(my_struc_Rventas(I).impuesto)
     
            'aqui es para el total
            ssuma5 = ssuma5 + Format(my_struc_Rventas(I).total)

        End If

        'fin 12/12/2017 pll
    Next I

    '*****************************
    'aqui los sub-totales
    h = 3
    v = v + 1
    'aqui los totales
    'provemos anteriormente tambien estaba asi
    objWorksheet.Cells(v, h + 6).Font.bold = True
    objWorksheet.Cells(v, h + 6).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 6) = "TOTALES"
   
    objWorksheet.Cells(v, h + 7).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 7) = "" & Format(ssuma1, "0.00") 'cantidad
    
    objWorksheet.Cells(v, h + 8).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 8) = "" & Format(ssuma2, "0.00") 'precio unitario
   
    objWorksheet.Cells(v, h + 9).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 9) = "" & Format(ssuma3, "0.00") ''vtaxproc.total
   
    objWorksheet.Cells(v, h + 10).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 10) = "'" & Format(ssuma4, "0.00") 'costo unitario impuesto
   
    objWorksheet.Cells(v, h + 11).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 11) = "'" & Format(ssuma5, "0.00") 'costo total
   
    objWorksheet.Cells(v, h + 12).Interior.color = RGB(215, 215, 0)  'resaltador
    objWorksheet.Cells(v, h + 12) = "" & Format(ssuma6, "0.00") 'ganancia
   
    Exit Sub
 
carga_Rcompras:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

    End Select

End Sub

'fin 13/12/2017 pll
