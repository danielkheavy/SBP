Attribute VB_Name = "Module5"

Public objExcel As Excel.Application

Public Function Inicio_Excel()

    Dim I As Integer

    Dim j As Integer

    On Error GoTo cmd78122_err

    Set objExcel = New Excel.Application
 
    objExcel.Visible = True 'lo hacemos visible
    objExcel.SheetsInNewWorkbook = 1 'decimos cuantas hojas queremos en el nuevo documento
    objExcel.Workbooks.Add '-- añadimos el objeto al workbook
    Inicio_Excel = 1
    Exit Function
cmd78122_err:
    Exit Function

End Function

Public Function Formato_Excel(Num_Campos As Integer, Nombre_Campos() As String) As Boolean

    With objExcel.ActiveSheet
        'MsgBox Num_Campos
        'Formato de las celdas de los titulos
        
        ''' kenyo 23/08/2017 Mejora reporte saldos a un periodo
        '.Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        '.Range(.Cells(3, 1), .Cells(3, 9)).Font.bold = True
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Font.bold = True
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Interior.color = RGB(192, 192, 250)
        ''' kenyo 23/08/2017 Mejora reporte saldos a un periodo
        
        For I = 1 To Num_Campos Step 1
            .Cells(3, I) = Nombre_Campos(I)
        Next I

        'hasta aki pa colocar los titulos
        'a partir de aki ta claro que es pa darle el ancho a las celdas ;-)
        
        ''''19/09/2017 kenyo Mejora Reporte Saldo Actual
               
        .columns("A").ColumnWidth = 11
        .columns("B").ColumnWidth = 12
        .columns("C").ColumnWidth = 35
        .columns("D").ColumnWidth = 7
        .columns("E").ColumnWidth = 6
        .columns("F").ColumnWidth = 11
        .columns("G").ColumnWidth = 11
        .columns("H").ColumnWidth = 11
        .columns("I").ColumnWidth = 11
        .columns("J").ColumnWidth = 11
        .columns("K").ColumnWidth = 8
        ''''19/09/2017 kenyo Mejora Reporte Saldo Actual
        
    End With

End Function

''' 01/11/2017 Mejora reporte lista de precios
Public Function Formato_ExcelListaPrecios(Num_Campos As Integer, _
                                          Nombre_Campos() As String) As Boolean

    With objExcel.ActiveSheet
        'MsgBox Num_Campos
        'Formato de las celdas de los titulos
        
        ''' kenyo 23/08/2017 Mejora reporte saldos a un periodo
        '.Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        '.Range(.Cells(3, 1), .Cells(3, 9)).Font.bold = True
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Font.bold = True
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Interior.color = RGB(192, 192, 250)
        ''' kenyo 23/08/2017 Mejora reporte saldos a un periodo
        
        For I = 1 To Num_Campos Step 1
            .Cells(3, I) = Nombre_Campos(I)
        Next I

        'hasta aki pa colocar los titulos
        'a partir de aki ta claro que es pa darle el ancho a las celdas ;-)
        
        ''''19/09/2017 kenyo Mejora Reporte Saldo Actual
        
        .columns("A").ColumnWidth = 11
        .columns("B").ColumnWidth = 11
        .columns("C").ColumnWidth = 30
        .columns("D").ColumnWidth = 6
        .columns("E").ColumnWidth = 10
        .columns("F").ColumnWidth = 6
        .columns("G").ColumnWidth = 10
        .columns("H").ColumnWidth = 6
        .columns("I").ColumnWidth = 10
        .columns("J").ColumnWidth = 10
        .columns("K").ColumnWidth = 10
        
        If Num_Campos <> "23" Then
            .columns("L").ColumnWidth = 15
            .columns("M").ColumnWidth = 15
            .columns("N").ColumnWidth = 15

        End If
       
        ''''19/09/2017 kenyo Mejora Reporte Saldo Actual
        
    End With

End Function

''' 01/11/2017 Mejora reporte lista de precios

Public Function Formato_SaldosPeriodo(Num_Campos As Integer, _
                                      Nombre_Campos() As String) As Boolean

    With objExcel.ActiveSheet
        
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Font.bold = True
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Interior.color = RGB(192, 192, 250)
         
        For I = 1 To Num_Campos Step 1
            .Cells(3, I) = Nombre_Campos(I)
        Next I
                     
        .columns("A").ColumnWidth = 11
        .columns("B").ColumnWidth = 35
        .columns("C").ColumnWidth = 7
        .columns("D").ColumnWidth = 7
        .columns("E").ColumnWidth = 8
        .columns("F").ColumnWidth = 10
        .columns("G").ColumnWidth = 10
        .columns("H").ColumnWidth = 10
        .columns("I").ColumnWidth = 10
        .columns("J").ColumnWidth = 10
        .columns("K").ColumnWidth = 8
        
    End With

End Function

Public Function Formato_Productosdiarios(Num_Campos As Integer, _
                                         Nombre_Campos() As String) As Boolean

    With objExcel.ActiveSheet
     
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Font.bold = True
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Interior.color = RGB(192, 192, 250)
     
        For I = 1 To Num_Campos Step 1
            .Cells(3, I) = Nombre_Campos(I)
        Next I
     
        .columns("A").ColumnWidth = 11
        .columns("B").ColumnWidth = 3
        .columns("C").ColumnWidth = 3
        .columns("D").ColumnWidth = 3
        .columns("E").ColumnWidth = 3
        .columns("F").ColumnWidth = 3
        .columns("G").ColumnWidth = 3
        .columns("H").ColumnWidth = 3
        .columns("I").ColumnWidth = 3
        .columns("J").ColumnWidth = 3
        
        .columns("K").ColumnWidth = 3
        .columns("L").ColumnWidth = 3
        .columns("M").ColumnWidth = 3
        .columns("N").ColumnWidth = 3
        .columns("O").ColumnWidth = 3
        
        .columns("P").ColumnWidth = 3
        .columns("Q").ColumnWidth = 3
        .columns("R").ColumnWidth = 3
        .columns("S").ColumnWidth = 3
                
        .columns("T").ColumnWidth = 3
        .columns("U").ColumnWidth = 3
        .columns("V").ColumnWidth = 3
        .columns("W").ColumnWidth = 3
        
        .columns("X").ColumnWidth = 3
        .columns("Y").ColumnWidth = 3
        .columns("Z").ColumnWidth = 3
        .columns("AA").ColumnWidth = 3
                
        .columns("AB").ColumnWidth = 3
        .columns("AC").ColumnWidth = 3
        .columns("AD").ColumnWidth = 3
        .columns("AE").ColumnWidth = 3
        .columns("AF").ColumnWidth = 3
        
    End With

End Function

'' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
'Public Function Formato_ExcelEntradasSalidas(Num_Campos As Integer, Nombre_Campos() As String) As Boolean
'With objExcel.ActiveSheet
'
'        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
'        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Font.bold = True
'        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Interior.color = RGB(192, 192, 250)
'
'
'    For I = 1 To Num_Campos Step 1
'        .Cells(3, I) = Nombre_Campos(I)
'    Next I
'
'        .columns("A").ColumnWidth = 9
'
'        If repinv.vesubfamilia = "S" Then
'        .columns("B").ColumnWidth = 9
'        Else
'        .columns("B").ColumnWidth = 0
'        End If
'
'        .columns("C").ColumnWidth = 13
'        .columns("D").ColumnWidth = 33
'        .columns("E").ColumnWidth = 7
'        .columns("F").ColumnWidth = 5
'        .columns("G").ColumnWidth = 9
'        .columns("H").ColumnWidth = 9
'        .columns("I").ColumnWidth = 9
'        .columns("J").ColumnWidth = 9
'        .columns("K").ColumnWidth = 9
'        .columns("L").ColumnWidth = 9
'        .columns("M").ColumnWidth = 9
'
'        .columns("N").ColumnWidth = 9
'        .columns("O").ColumnWidth = 9
'
'End With
'End Function
Public Function Formato_ExcelEntradasSalidas(Num_Campos As Integer, _
                                             Nombre_Campos() As String) As Boolean

    With objExcel.ActiveSheet
    
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Font.bold = True
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Interior.color = RGB(192, 192, 250)
        
        For I = 1 To Num_Campos Step 1
            .Cells(3, I) = Nombre_Campos(I)
        Next I
                      
        .columns("A").ColumnWidth = 9
        
        If repinv.vesubfamilia = "S" Then
            .columns("B").ColumnWidth = 9
        Else
            .columns("B").ColumnWidth = 0

        End If
        
        .columns("C").ColumnWidth = 12
        .columns("D").ColumnWidth = 29
        .columns("E").ColumnWidth = 7
        .columns("F").ColumnWidth = 5
        
        If repinv.ChkSaldoInicial.Value = 1 Then
            .columns("G").ColumnWidth = 10
        Else
            .columns("G").ColumnWidth = 0

        End If
        
        .columns("H").ColumnWidth = 8
        .columns("I").ColumnWidth = 8
        .columns("J").ColumnWidth = 8
        .columns("K").ColumnWidth = 8
        .columns("L").ColumnWidth = 9
        .columns("M").ColumnWidth = 9
        .columns("N").ColumnWidth = 8
        .columns("O").ColumnWidth = 8
        .columns("P").ColumnWidth = 8
        
    End With

End Function

'' 16/01/2018 Stock Inicial en Kardex Producto y Sunat

''''13/09/2017 kenyo Reporte Comisiones Productos en Excel
Public Function Formato_ExcelComision(Num_Campos As Integer, _
                                      Nombre_Campos() As String) As Boolean

    With objExcel.ActiveSheet
        'MsgBox Num_Campos
        'Formato de las celdas de los titulos
        
        ''' kenyo 23/08/2017 Mejora reporte saldos a un periodo
        '.Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        '.Range(.Cells(3, 1), .Cells(3, 9)).Font.bold = True
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Font.bold = True
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Interior.color = RGB(192, 192, 250)
        ''' kenyo 23/08/2017 Mejora reporte saldos a un periodo
        
        For I = 1 To Num_Campos Step 1
            .Cells(3, I) = Nombre_Campos(I)
        Next I

        'hasta aki pa colocar los titulos
        'a partir de aki ta claro que es pa darle el ancho a las celdas ;-)
        .columns("A").ColumnWidth = 11
        .columns("B").ColumnWidth = 5
        .columns("C").ColumnWidth = 7
        .columns("D").ColumnWidth = 10
        .columns("E").ColumnWidth = 10
        .columns("F").ColumnWidth = 10
        .columns("G").ColumnWidth = 35.5
        .columns("H").ColumnWidth = 7
        .columns("I").ColumnWidth = 5
        .columns("J").ColumnWidth = 10
        .columns("K").ColumnWidth = 10
        .columns("L").ColumnWidth = 10
        .columns("M").ColumnWidth = 5
        .columns("N").ColumnWidth = 7
        .columns("O").ColumnWidth = 5
        .columns("P").ColumnWidth = 5
        .columns("Q").ColumnWidth = 7
        .columns("R").ColumnWidth = 10
        .columns("S").ColumnWidth = 7
        .columns("T").ColumnWidth = 12
        
    End With

End Function

''''13/09/2017 kenyo Reporte Comisiones Productos en Excel

' ''10/10/2017 Reporte de Seguimiento de facturas En Excel
Public Function Formato_ExcelSeguimiento(Num_Campos As Integer, _
                                         Nombre_Campos() As String) As Boolean

    With objExcel.ActiveSheet

        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Font.bold = True
        
        If repdocum.vdetalle = "S" Then
            .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Interior.color = RGB(232, 232, 232)
        Else
            .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Interior.color = RGB(192, 192, 250)

        End If
      
        For I = 1 To Num_Campos Step 1
            .Cells(3, I) = Nombre_Campos(I)
        Next I
        
        .columns("A").ColumnWidth = 11
        .columns("B").ColumnWidth = 11
        
        .columns("C").ColumnWidth = 4
        .columns("D").ColumnWidth = 7
        .columns("E").ColumnWidth = 10
        .columns("F").ColumnWidth = 14
        .columns("G").ColumnWidth = 40
        .columns("H").ColumnWidth = 4
        
        .columns("I").ColumnWidth = 10
        .columns("J").ColumnWidth = 10
        .columns("K").ColumnWidth = 10
        .columns("L").ColumnWidth = 10
        .columns("M").ColumnWidth = 7
        .columns("N").ColumnWidth = 7
        .columns("O").ColumnWidth = 7
        .columns("P").ColumnWidth = 7
        .columns("Q").ColumnWidth = 7
        .columns("R").ColumnWidth = 7
        .columns("S").ColumnWidth = 7
        
        '08/05/2018 Reporte Pedidos Orden de Trabajo en Excel
        .columns("T").ColumnWidth = 15
        '08/05/2018 Reporte Pedidos Orden de Trabajo en Excel
        
        ''02/11/2017 Reporte de Seguimiento de facturas incluye delivery

        If repdocum.vfpago = "S" And repdocum.vedelivery = "S" Then
            .columns("U").ColumnWidth = 10
            .columns("V").ColumnWidth = 10
            .columns("W").ColumnWidth = 10
            .columns("X").ColumnWidth = 10
            .columns("Y").ColumnWidth = 10
            .columns("Z").ColumnWidth = 10
             
            .columns("AB").ColumnWidth = 10
            .columns("AC").ColumnWidth = 25
            .columns("AD").ColumnWidth = 25

        End If
        
        If repdocum.vfpago = "S" And repdocum.vedelivery = "N" Then
            .columns("U").ColumnWidth = 10
            .columns("V").ColumnWidth = 10
            .columns("W").ColumnWidth = 10
            .columns("X").ColumnWidth = 10
            .columns("Y").ColumnWidth = 10
            .columns("Z").ColumnWidth = 10

        End If
           
        If repdocum.vfpago = "N" And repdocum.vedelivery = "S" Then
            'SOLO DATOS DE DELIVERY
            .columns("U").ColumnWidth = 10
            .columns("V").ColumnWidth = 25
            .columns("W").ColumnWidth = 25
        
        End If
        
        ''02/11/2017 Reporte de Seguimiento de facturas incluye delivery
                
    End With

End Function

Public Function Formato_ExcelSeguimientoDetalle(Num_Campos As Integer, _
                                                Nombre_Campos() As String) As Boolean

    With objExcel.ActiveSheet

        .Range(.Cells(4, 6), .Cells(4, Num_Campos)).Font.bold = True
        .Range(.Cells(4, 6), .Cells(4, Num_Campos)).Borders.LineStyle = xlContinuous
        '.Range(.Cells(4, 5), .Cells(4, Num_Campos)).Interior.color = RGB(192, 192, 250)
      
        For I = 6 To 13
            .Cells(4, I) = Nombre_Campos(I)
        Next I

        .columns("F").ColumnWidth = 14
        .columns("G").ColumnWidth = 40
        .columns("H").ColumnWidth = 4
        .columns("I").ColumnWidth = 10
        .columns("J").ColumnWidth = 10
        .columns("K").ColumnWidth = 10
        .columns("L").ColumnWidth = 10

    End With

End Function

' ''10/10/2017 Reporte de Seguimiento de facturas En Excel

''''09/10/2017 kenyo Testing Reportes
Public Function Formato_Excel2(Num_Campos As Integer, _
                               Nombre_Campos() As String) As Boolean

    With objExcel.ActiveSheet
      
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Font.bold = True
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Interior.color = RGB(192, 192, 250)
        
        For I = 1 To Num_Campos Step 1
            .Cells(3, I) = Nombre_Campos(I)
        Next I
                   
        .columns("A").ColumnWidth = 12
        .columns("B").ColumnWidth = 35
        .columns("C").ColumnWidth = 10
        .columns("D").ColumnWidth = 6
        .columns("E").ColumnWidth = 10
        .columns("F").ColumnWidth = 10
        .columns("G").ColumnWidth = 10
        .columns("H").ColumnWidth = 10
        .columns("I").ColumnWidth = 10
        .columns("J").ColumnWidth = 10
        .columns("K").ColumnWidth = 8
        
    End With

End Function

''''09/10/2017 kenyo Testing Reportes

'' 11/12/2017 SubReceta
Public Function Formato_ExcelReceta(Num_Campos As Integer, _
                                    Nombre_Campos() As String) As Boolean

    With objExcel.ActiveSheet
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Font.bold = True
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Interior.color = RGB(192, 192, 250)
            
        For I = 1 To Num_Campos Step 1
            .Cells(3, I) = Nombre_Campos(I)
        Next I
            
        .columns("A").ColumnWidth = 11
        .columns("B").ColumnWidth = 35
        .columns("C").ColumnWidth = 9
        .columns("D").ColumnWidth = 8
        .columns("E").ColumnWidth = 8
        .columns("F").ColumnWidth = 9
        .columns("G").ColumnWidth = 9
        
    End With

End Function

'' 11/12/2017 SubReceta

