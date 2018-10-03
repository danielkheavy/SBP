Attribute VB_Name = "backup_tablas"

'inicio 25/04/2017 pll
'aqui explicacion tecnica para este script que lo veras varias veces
' hastaCuanto = 11 - Len(mytablef.Fields("VENDEDOR")) 'hasta cuanto llenara los espacios en blanco en funcion a lo llenado
' Call llenar_datos(hastaCuanto, mydato, nuevoDato) 'esta funcion los llena los espacios en blanco
'  myREG = myREG & nuevoDato ' y aqui imprime
' myREG = myREG & "&"
Type struc_correo

    myserver                              As String
    myusername                            As String
    mypassword                            As String
    myfromname                            As String
    myfromemail                           As String
    myport                                As String
    myselecciona                          As String
    mytto                                 As String
    mychkssl                              As String
    myattach                              As String
    mysubjec                              As String
    mymsg                                 As String
    myhtml                                As String

End Type

Global mystruc_correo As struc_correo

Public myfile         As String

Public origen         As String

Public creafilerar    As String

Public Sub bkp_detalle(fechai As String, fechaf As String, conta_record As Integer)

    Dim mysql       As String

    Dim mytablef    As New ADODB.Recordset

    Dim hastaCuanto As Integer

    Dim myDato      As String

    Dim nuevoDato   As String

    FileName = "C:\EmpaquetaVi\" & "bkpdetalle" & ".txt"

    mysql = "SELECT isnull(tipo,'vac') as tipo," & Chr$(10)
    mysql = mysql & "isnull(SERIE,'vaci') as SERIE," & Chr$(10)
    mysql = mysql & "isnull(NUMERO,'vacio') as NUMERO," & Chr$(10)
    mysql = mysql & "isnull(tipoclie,'v') as tipoclie," & Chr$(10)
    mysql = mysql & "isnull(codigo,'vacio') as codigo," & Chr$(10)
    mysql = mysql & "isnull(ACU,'?') as ACU," & Chr$(10)
    mysql = mysql & "isnull(ACU1,'?') as ACU1," & Chr$(10)
    mysql = mysql & "isnull(fecha,0) as fecha , " & Chr$(10)
    mysql = mysql & "isnull(moneda,'?') as moneda, " & Chr$(10)
    mysql = mysql & "isnull(producto,'vacio') as producto, " & Chr$(10)
    mysql = mysql & "isnull(DESCRIPCIO,'vacio') as DESCRIPCIO," & Chr$(10)
    mysql = mysql & "isnull(unidad,'vacio') as unidad , " & Chr$(10)
    mysql = mysql & "isnull(factor,0) as factor, " & Chr$(10)
    mysql = mysql & "isnull(cantidad,0) as cantidad, " & Chr$(10)
    mysql = mysql & "isnull(PRECIO,0) as PRECIO," & Chr$(10)
    mysql = mysql & "isnull(igv,0) as igv , " & Chr$(10)
    mysql = mysql & "isnull(neto,0) as neto, " & Chr$(10)
    mysql = mysql & "isnull(descuento,0) as descuento, " & Chr$(10)
    mysql = mysql & "isnull(subtotal,0) as subtotal , " & Chr$(10)
    mysql = mysql & "isnull(IMPUESTO,0) as IMPUESTO," & Chr$(10)
    mysql = mysql & "isnull(total,0) as total , " & Chr$(10)
    mysql = mysql & "isnull(estado,'v') as estado, " & Chr$(10)
    mysql = mysql & "isnull(usuario,'vacio') as usuario, " & Chr$(10)
    mysql = mysql & "isnull(FECHACREA,0) as FECHACREA, " & Chr$(10)
    mysql = mysql & "isnull(HORA,0) as HORA," & Chr$(10)
    mysql = mysql & "isnull(vendedor,'vacio') as VENDEDOR," & Chr$(10)
    mysql = mysql & "isnull(BODEGA,'va') as BODEGA," & Chr$(10) & Chr$(10)
    mysql = mysql & "isnull(BODEGAF,'va') as BODEGAF," & Chr$(10) & Chr$(10)
    mysql = mysql & "isnull(DESLIPO,0) as DESLIPO," & Chr$(10) & Chr$(10)
    mysql = mysql & "isnull(FLAGE,'¿') as FLAGE," & Chr$(10) & Chr$(10)
    mysql = mysql & "isnull(LINEA,'vacio') as LINEA," & Chr$(10)
    mysql = mysql & "isnull(T1,0) as T1," & Chr$(10)
    mysql = mysql & "isnull(T2,0) as T2," & Chr$(10)
    mysql = mysql & "isnull(T3,0) as T3," & Chr$(10)
    mysql = mysql & "isnull(T4,0) as T4," & Chr$(10)
    mysql = mysql & "isnull(t5,0) as t5," & Chr$(10)
    mysql = mysql & "isnull(t6,0) as t6," & Chr$(10)
    mysql = mysql & "isnull(t7,0) as t7," & Chr$(10)
    mysql = mysql & "isnull(t8,0) as t8," & Chr$(10)
    mysql = mysql & "isnull(t9,0) as t9," & Chr$(10)
    mysql = mysql & "isnull(t10,0) as t10," & Chr$(10)
    mysql = mysql & "isnull(t11,0) as t11," & Chr$(10)
    mysql = mysql & "isnull(t12,0) as t12," & Chr$(10)
    mysql = mysql & "isnull(t13,0) as t13, " & Chr$(10)
    mysql = mysql & "isnull(t14,0) as t14," & Chr$(10)
    mysql = mysql & "isnull(t15,0) as t15," & Chr$(10)
    mysql = mysql & "isnull(t16,0) as t16," & Chr$(10)
    mysql = mysql & "isnull(L1,'?') as L1," & Chr$(10)
    mysql = mysql & "isnull(L2,'?') as L2," & Chr$(10)
    mysql = mysql & "isnull(L3,'?') as L3," & Chr$(10)
    mysql = mysql & "isnull(L4,'?') as L4," & Chr$(10)
    mysql = mysql & "isnull(LOCAL,'vacio') as LOCAL," & Chr$(10)
    mysql = mysql & "isnull(PROVEEDORP,'vacio') as PROVEEDORP," & Chr$(10)
    mysql = mysql & "isnull(OBSERVA1,'vacio') as OBSERVA1," & Chr$(10)
    mysql = mysql & "isnull(OBSERVA2,'vacio') as OBSERVA2," & Chr$(10)
    mysql = mysql & "isnull(OBSERVA3,'vacio') as OBSERVA3," & Chr$(10)
    mysql = mysql & "isnull(OBSERVA4,'vacio') as OBSERVA4," & Chr$(10)
    mysql = mysql & "isnull(ZONA,'vacio') as ZONA," & Chr$(10)
    mysql = mysql & "isnull(ISC,0) as ISC," & Chr$(10)
    mysql = mysql & "isnull(TAX,0) as TAX," & Chr$(10)
    mysql = mysql & "isnull(VTANETA,0) as VTANETA," & Chr$(10)
    mysql = mysql & "isnull(TCOSTO,0) as TCOSTO," & Chr$(10)
    mysql = mysql & "isnull(ganancia,0) as ganancia , " & Chr$(10)
    mysql = mysql & "isnull(comision,0) as comision, " & Chr$(10)
    mysql = mysql & "isnull(cajero,'vacio') as cajero, " & Chr$(10)
    mysql = mysql & "isnull(caja,'vacio') as caja, " & Chr$(10)
    mysql = mysql & "isnull(turno,'?') as turno, " & Chr$(10)
    mysql = mysql & "isnull(servicio,'?') as servicio," & Chr$(10)
    mysql = mysql & "isnull(comanda,'vacio') as comanda," & Chr$(10)
    mysql = mysql & "isnull(MESA,'vacio') as MESA," & Chr$(10)
    mysql = mysql & "isnull(SALON,'vacio') as SALON," & Chr$(10)
    mysql = mysql & "isnull(MESERO,'vacio') as MESERO," & Chr$(10)
    mysql = mysql & "isnull(SENTIDO,'?') as SENTIDO," & Chr$(10)
    mysql = mysql & "isnull(CCOSTO,'vacio') as CCOSTO,"
    mysql = mysql & "isnull(familia,'vacio') as familia , " & Chr$(10)
    mysql = mysql & "isnull(subfamilia,'vacio') as subfamilia , " & Chr$(10)
    mysql = mysql & "isnull(marca,'vacio') as marca, " & Chr$(10)
    mysql = mysql & "isnull(percepcion,0) as percepcion, " & Chr$(10)
    mysql = mysql & "isnull(TPERCEPCIO,0) as TPERCEPCIO, " & Chr$(10)
    mysql = mysql & "isnull(FLETE,0) as FLETE," & Chr$(10)
    mysql = mysql & "isnull(LOCALF,'vac') as LOCALF," & Chr$(10)
    mysql = mysql & "isnull(IVAP,0) as IVAP," & Chr$(10)
    mysql = mysql & "isnull(TIVAP,0) as TIVAP," & Chr$(10)
    mysql = mysql & "isnull(NROPRECIO,'?') as NROPRECIO," & Chr$(10)
    mysql = mysql & "isnull(TISC,0) as TISC," & Chr$(10)
    mysql = mysql & "isnull(PLACA,'vacio') as PLACA," & Chr$(10)
    mysql = mysql & "isnull(xneto,0) as xneto , " & Chr$(10)
    mysql = mysql & "isnull(tdetra,0) as tdetra, " & Chr$(10)
    mysql = mysql & "isnull(DENUMERO,'vacio') as DENUMERO, " & Chr$(10)
    mysql = mysql & "isnull(categoria,'vacio') as categoria, " & Chr$(10)
    mysql = mysql & "isnull(ADUANA,'vacio') as ADUANA," & Chr$(10)
    mysql = mysql & "isnull(DUA,'vacio') as DUA," & Chr$(10)
    mysql = mysql & "isnull(cantdev,0) as cantdev," & Chr$(10)
    mysql = mysql & "isnull(servicioco,0) as servicioco , " & Chr$(10)
    mysql = mysql & "isnull(serviciopo,0) as serviciopo, " & Chr$(10)
    mysql = mysql & "isnull(destopo,0) as destopo, " & Chr$(10)
    mysql = mysql & "isnull(convert(varchar,fechaborra,108),'0') as fechaborra," & Chr$(10)
    mysql = mysql & "isnull(horaborra,'vacio') as horaborra , " & Chr$(10)
    mysql = mysql & "isnull(administrador,'vacio') as administrador, " & Chr$(10)
    mysql = mysql & "isnull(detraccion,0) as detraccion " & Chr$(10)
    mysql = mysql & "from detalle where fecha>='" & Format(fechai, "YYYYMMDD") & "'" & Chr$(10)
    mysql = mysql & "and fecha<='" & Format(fechaf, "YYYYMMDD") & "' " & Chr$(10)
    mysql = mysql & "order by fecha"

    mytablef.Open mysql, cn, adOpenStatic, adLockOptimistic
  
    If mytablef.RecordCount > 0 Then
        Do

            If mytablef.EOF Then Exit Do
            'aqui es para abrir el file txt y llenarlo
            Filelibero1 = FreeFile
            Open FileName For Append As #Filelibero1
            conta = conta + 1
      
            conta_record = mytablef.RecordCount
            myREG = ""
      
            If mytablef.Fields("tipo") = "vac" Then
                myREG = myREG & Space$(3) & "&"
            Else
                hastaCuanto = 3 - Trim(Len(mytablef.Fields("tipo")))
                myDato = mytablef.Fields("tipo")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("SERIE") = "vaci" Then
                myREG = myREG & Space$(4)
                myREG = myREG & "&"
            Else
                hastaCuanto = 4 - Len(mytablef.Fields("SERIE"))
                myDato = mytablef.Fields("SERIE")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'SERIE 2
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("NUMERO")) = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("NUMERO"))
                myDato = mytablef.Fields("NUMERO")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'NUMERO 3
                myREG = myREG & nuevoDato
                myREG = myREG & "&" 'gion separador

            End If

            If mytablef.Fields("tipoclie") = "v" Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                hastaCuanto = 1 - Len(mytablef.Fields("tipoclie"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("tipoclie"), nuevoDato) 'tipoclie 4
                myREG = myREG & nuevoDato
                myREG = myREG & "&" 'gion separador

            End If

            If mytablef.Fields("codigo") = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("codigo"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("codigo"), nuevoDato) 'codigo 5
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("ACU") = "?" Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                hastaCuanto = 1 - Len(mytablef.Fields("ACU"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("ACU"), nuevoDato) 'ACU 6
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("ACU1") = "?" Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                hastaCuanto = 1 - Len(mytablef.Fields("ACU1"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("ACU1"), nuevoDato) 'ACU1 7
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("FECHA")) = 0 Then
                myREG = myREG & Space$(10)
                myREG = myREG & "&"
            Else
                hastaCuanto = 1 - Len(mytablef.Fields("FECHA"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("FECHA"), nuevoDato) 'FECHA 8
                myREG = myREG & nuevoDato
                myREG = myREG & "&" 'gion separador

            End If

            If Trim(mytablef.Fields("MONEDA")) = "?" Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                hastaCuanto = 1 - Len(mytablef.Fields("MONEDA"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("MONEDA"), nuevoDato) 'MONEDA 9
                myREG = myREG & nuevoDato
                myREG = myREG & "&" 'gion separador

            End If

            If Trim(mytablef.Fields("PRODUCTO")) = "vacio" Then
                myREG = myREG & Space$(15)
                myREG = myREG & "&"
            Else
                hastaCuanto = 15 - Len(mytablef.Fields("PRODUCTO"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("PRODUCTO"), nuevoDato) 'PRODUCTO 10
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("DESCRIPCIO")) = "vacio" Then
                myREG = myREG & Space$(120)
                myREG = myREG & "&"
            Else
                hastaCuanto = 120 - Len(mytablef.Fields("DESCRIPCIO"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("DESCRIPCIO"), nuevoDato) 'DESCRIPCIO 11
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("UNIDAD")) = "vacio" Then
                myREG = myREG & Space$(6)
                myREG = myREG & "&"
            Else
                hastaCuanto = 6 - Len(mytablef.Fields("UNIDAD"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("UNIDAD"), nuevoDato)      'UNIDAD 12
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("FACTOR") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("FACTOR"))
                myDato = mytablef.Fields("FACTOR")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)    'FACTOR 13
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("cantidad") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Trim(Len(mytablef.Fields("cantidad")))
                myDato = Trim(mytablef.Fields("cantidad"))
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'cantidad 14
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("PRECIO") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Trim(Len(mytablef.Fields("PRECIO")))
                myDato = Trim(mytablef.Fields("PRECIO"))
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'PRECIO 15
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("IGV") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("IGV"))
                myDato = mytablef.Fields("IGV")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'igv 16
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("NETO") = 0 Then
                myREG = myREG & Space$(10)
                myREG = myREG & "&"
            Else
                hastaCuanto = 10 - Len(mytablef.Fields("NETO"))
                myDato = mytablef.Fields("NETO")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)  'NETO 17
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("DESCUENTO") = 0 Then
                myREG = myREG & Space$(10)
                myREG = myREG & "&"
            Else
                hastaCuanto = 10 - Len(mytablef.Fields("DESCUENTO"))      'DESCUENTO 18
                myDato = mytablef.Fields("DESCUENTO")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("SUBTOTAL") = 0 Then
                myREG = myREG & Space$(11) & (0)
                myREG = myREG & "&"
            Else
                hastaCuanto = 12 - Len(mytablef.Fields("SUBTOTAL")) 'SUBTOTAL 19
                myDato = mytablef.Fields("SUBTOTAL")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("IMPUESTO") = 0 Then
                myREG = myREG & Space$(11) & (0)
                myREG = myREG & "&"
            Else
                hastaCuanto = 12 - Len(mytablef.Fields("IMPUESTO"))
                myDato = mytablef.Fields("IMPUESTO")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'IMPUESTO 20
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("TOTAL") = 0 Then
                myREG = myREG & Space$(11) & (0)
                myREG = myREG & "&"
            Else
                hastaCuanto = 12 - Len(mytablef.Fields("TOTAL"))
                myDato = mytablef.Fields("TOTAL")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'TOTAL 21
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("ESTADO") = "v" Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                myREG = myREG & mytablef.Fields("ESTADO")         'ESTADO 22
                myREG = myREG & "&" 'gion separador

            End If

            If mytablef.Fields("USUARIO") = "vacio" Then
                myREG = myREG & Space$(12)
                myREG = myREG & "&"
            Else
                hastaCuanto = 12 - Len(mytablef.Fields("USUARIO"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("USUARIO"), nuevoDato) 'USUARIO 23
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("FECHACREA") = 0 Then
                myREG = myREG & Space$(10)
                myREG = myREG & "&"
            Else
                myREG = myREG & mytablef.Fields("FECHACREA")  'FECHACREA 24
                myREG = myREG & "&" 'gion separador

            End If

            If mytablef.Fields("HORA") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                myREG = myREG & mytablef.Fields("HORA")                                'HORA 25
                myREG = myREG & "&" 'gion separador

            End If

            If Trim(mytablef.Fields("VENDEDOR")) = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("VENDEDOR"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("VENDEDOR"), nuevoDato) 'VENDEDOR 26
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("BODEGA") = "va" Then
                myREG = myREG & Space$(2)
                myREG = myREG & "&"
            Else
                myREG = myREG & mytablef.Fields("BODEGA")                              'BODEGA 27
                myREG = myREG & "&" 'gion separador

            End If

            If mytablef.Fields("BODEGAF") = "va" Then
                myREG = myREG & Space$(2)
                myREG = myREG & "&"
            Else
                hastaCuanto = 2 - Len(mytablef.Fields("BODEGAF"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("BODEGAF"), nuevoDato) 'BODEGAF 28
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("DESLIPO") = 0 Then
                myREG = myREG & Space$(10)
                myREG = myREG & "&"
            Else
                hastaCuanto = 10 - Len(mytablef.Fields("DESLIPO"))
                myDato = mytablef.Fields("DESLIPO")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'DESLIPO 29
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("FLAGE") = "" Or mytablef.Fields("FLAGE") = "¿" Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                myREG = myREG & mytablef.Fields("FLAGE")                                'FLAGE 30
                myREG = myREG & "&"

            End If

            If mytablef.Fields("LINEA") = "vacio" Then
                myREG = myREG & Space$(6) & "&"
            Else
                hastaCuanto = 6 - Len(mytablef.Fields("LINEA"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("LINEA"), nuevoDato)  'LINEA 31
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("T1") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("T1"))
                myDato = mytablef.Fields("T1")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'T1 32
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("T2") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("T2"))
                myDato = mytablef.Fields("T2")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'T2 33
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("T3") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("T3"))
                myDato = mytablef.Fields("T3")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'T3 34
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("T4") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("T4"))
                myDato = mytablef.Fields("T4")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'T4 35
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("T5") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("T5"))
                myDato = mytablef.Fields("T5")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'T5 36
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("T6") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("T6"))
                myDato = mytablef.Fields("T6")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'T6 37
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("T7") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("T7"))
                myDato = mytablef.Fields("T7")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'T7 38
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("T8") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("T8"))
                myDato = mytablef.Fields("T8")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'T8 39
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("T9") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("T9"))
                myDato = mytablef.Fields("T9")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'T9 40
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("T10") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("T10"))
                myDato = mytablef.Fields("T10")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'T10 41
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("T11") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("T11"))
                myDato = mytablef.Fields("T11")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'T11 42
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("T12") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("T12"))
                myDato = mytablef.Fields("T12")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'T12 43
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("T13") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("T13"))
                myDato = mytablef.Fields("T13")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'T13 44
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("T14") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("T14"))
                myDato = mytablef.Fields("T14")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'T14 45
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("T15") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("T15"))
                myDato = mytablef.Fields("T15")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'T15 46
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("T16") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("T16"))
                myDato = mytablef.Fields("T16")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'T16 47
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("L1") = "" Or mytablef.Fields("L1") = "?" Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                myREG = myREG & mytablef.Fields("L1")           'L1 48
                myREG = myREG & "&"

            End If

            If mytablef.Fields("L2") = "?" Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                myREG = myREG & mytablef.Fields("L2")           'L2 49
                myREG = myREG & "&"

            End If

            If mytablef.Fields("L3") = "?" Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                myREG = myREG & mytablef.Fields("L3")         'L3 50
                myREG = myREG & "&"

            End If

            If mytablef.Fields("L4") = "?" Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                myREG = myREG & mytablef.Fields("L4")         'L4 51
                myREG = myREG & "&"

            End If

            If mytablef.Fields("LOCAL") = "vacio" Then
                yREG = myREG & Space$(6)
                myREG = myREG & "&"
            Else
                hastaCuanto = 6 - Len(mytablef.Fields("LOCAL"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("LOCAL"), nuevoDato) 'LOCAL  52
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("PROVEEDORP") = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("PROVEEDORP"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("PROVEEDORP"), nuevoDato) 'PROVEEDORP 53
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("OBSERVA1") = "vacio" Then
                myREG = myREG & Space(150)
                myREG = myREG & "&"
            Else
                hastaCuanto = 150 - Len(mytablef.Fields("OBSERVA1"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("OBSERVA1"), nuevoDato) 'OBSERVA1 54
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("OBSERVA2") = "vacio" Then
                myREG = myREG & Space(150)
                myREG = myREG & "&"
            Else
                hastaCuanto = 150 - Len(mytablef.Fields("OBSERVA2"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("OBSERVA2"), nuevoDato) 'OBSERVA2 55
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("OBSERVA3") = "vacio" Then
                myREG = myREG & Space(150)
                myREG = myREG & "&"
            Else
                hastaCuanto = 150 - Len(mytablef.Fields("OBSERVA3"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("OBSERVA3"), nuevoDato) 'OBSERVA3 56
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("OBSERVA4") = "vacio" Then
                myREG = myREG & Space(150)
                myREG = myREG & "&"
            Else
                hastaCuanto = 150 - Len(mytablef.Fields("OBSERVA4"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("OBSERVA4"), nuevoDato) 'OBSERVA4 57
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("zona") = "vacio" Then
                myREG = myREG & Space$(6)
                myREG = myREG & "&"
            Else
                hastaCuanto = 6 - Len(mytablef.Fields("zona"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("zona"), nuevoDato) 'zona 58
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("ISC") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("ISC"))
                myDato = mytablef.Fields("ISC")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'ISC 59
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("TAX") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("TAX"))
                myDato = mytablef.Fields("TAX")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'TAX 60
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("VTANETA") = 0 Then
                myREG = myREG & Space$(7) & 0
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("VTANETA"))
                myDato = mytablef.Fields("VTANETA")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'VTANETA 61
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("TCOSTO") = 0 Then
                myREG = myREG & Space$(7) & 0
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("TCOSTO"))
                myDato = mytablef.Fields("TCOSTO")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'TCOSTO 62
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("GANANCIA") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("GANANCIA"))
                myDato = mytablef.Fields("GANANCIA")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'GANANCIA 63
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("COMISION") = 0 Then
                myREG = myREG & Space$(10)
                myREG = myREG & "&"
            Else
                hastaCuanto = 10 - Len(mytablef.Fields("COMISION"))
                myDato = mytablef.Fields("COMISION")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'COMISION 64
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("CAJERO") = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("CAJERO"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("CAJERO"), nuevoDato) 'CAJERO 65
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("CAJA") = "vacio" Then
                myREG = myREG & Space$(3)
                myREG = myREG & "&"
            Else
                hastaCuanto = 3 - Len(mytablef.Fields("CAJA"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("CAJA"), nuevoDato) 'CAJA 66
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Len(mytablef.Fields("TURNO")) = 0 Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                myREG = myREG & mytablef.Fields("TURNO")        'TURNO 67
                myREG = myREG & "&" 'gion separador

            End If

            If mytablef.Fields("SERVICIO") = "?" Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                hastaCuanto = 1 - Len(mytablef.Fields("SERVICIO"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("SERVICIO"), nuevoDato) 'SERVICIO 68
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("COMANDA") = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("COMANDA"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("COMANDA"), nuevoDato) 'COMANDA 69
                myREG = myREG & nuevoDato & ""
                myREG = myREG & "&"

            End If

            If mytablef.Fields("MESA") = "vac" Then
                myREG = myREG & Space$(3)
                myREG = myREG & "&"
            Else
                hastaCuanto = 3 - Len(mytablef.Fields("MESA"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("MESA"), nuevoDato) 'MESA 70
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("SALON") = "vac" Then
                myREG = myREG & Space$(3)
                myREG = myREG & "&"
            Else
                hastaCuanto = 3 - Len(mytablef.Fields("SALON"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("SALON"), nuevoDato) 'SALON 71
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("MESERO") = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("MESERO"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("MESERO"), nuevoDato) 'MESERO 72
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("SENTIDO") = "?" Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                hastaCuanto = 1 - Len(mytablef.Fields("SENTIDO"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("SENTIDO"), nuevoDato) 'SENTIDO 73
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("CCOSTO") = "vacio" Then
                myREG = myREG & Space(6)
                myREG = myREG & "&"
            Else
                hastaCuanto = 6 - Len(mytablef.Fields("CCOSTO"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("CCOSTO"), nuevoDato) 'CCOSTO 74
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("FAMILIA")) = "vacio" Then
                myREG = myREG & Space(6)
                myREG = myREG & "&"
            Else
                hastaCuanto = 6 - Len(mytablef.Fields("FAMILIA"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("FAMILIA"), nuevoDato) 'FAMILIA 75
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("SUBFAMILIA") = "vacio" Then
                myREG = myREG & Space$(6)
                myREG = myREG & "&"
            Else
                hastaCuanto = 6 - Len(mytablef.Fields("SUBFAMILIA"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("SUBFAMILIA"), nuevoDato) 'SUBFAMILIA 76
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("MARCA") = "vacio" Then
                myREG = myREG & Space$(6)
                myREG = myREG & "&"
            Else
                hastaCuanto = 6 - Len(mytablef.Fields("MARCA"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("MARCA"), nuevoDato) 'MARCA 77
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("PERCEPCION") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("PERCEPCION"))
                nuevoDato = mytablef.Fields("PERCEPCION")
                Call llenar_datos(hastaCuanto, nuevoDato, nuevoDato) 'PERCEPCION 78
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("TPERCEPCIO") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("TPERCEPCIO"))
                myDato = mytablef.Fields("TPERCEPCIO")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'TPERCEPCIO 79
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("FLETE") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("FLETE"))
                myDato = mytablef.Fields("FLETE")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'FLETE 80
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("LOCALF")) = "vac" Then
                myREG = myREG & Space$(3)
                myREG = myREG & "&"
            Else
                hastaCuanto = 3 - Len(mytablef.Fields("LOCALF"))
                myDato = Trim$(mytablef.Fields("LOCALF"))
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'LOCALF 81
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("IVAP") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("IVAP"))
                myDato = mytablef.Fields("IVAP")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) ' IVAP 82
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("TIVAP") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("TIVAP"))
                myDato = mytablef.Fields("TIVAP")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'TIVAP 83
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("NROPRECIO") = "?" Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                hastaCuanto = 1 - Len(mytablef.Fields("NROPRECIO"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("NROPRECIO"), nuevoDato) 'NROPRECIO 84
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("TISC") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("TISC"))
                myDato = mytablef.Fields("TISC")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'TISC 85
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("PLACA") = "vacio" Then
                myREG = myREG & Space(15)
                myREG = myREG & "&"
            Else
                hastaCuanto = 15 - Len(mytablef.Fields("PLACA"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("PLACA"), nuevoDato) ' PLACA 86
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("XNETO") = 0 Then
                myREG = myREG & Space$(7) & 0
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("XNETO"))
                myDato = mytablef.Fields("XNETO")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'XNETO 87
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("TDETRA") = 0 Then
                myREG = myREG & Space$(7) & 0
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("TDETRA"))
                myDato = mytablef.Fields("TDETRA")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'TDETRA 88
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("DENUMERO") = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("DENUMERO"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("DENUMERO"), nuevoDato) 'DENUMERO 89
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("CATEGORIA") = "vacio" Then
                myREG = myREG & Space$(6)
                myREG = myREG & "&"
            Else
                hastaCuanto = 6 - Len(mytablef.Fields("CATEGORIA"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("CATEGORIA"), nuevoDato) 'CATEGORIA 90
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("ADUANA") = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("ADUANA"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("ADUANA"), nuevoDato) 'ADUANA 91
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("DUA") = "vacio" Then
                myREG = myREG & Space$(10)
                myREG = myREG & "&"
            Else
                hastaCuanto = 10 - Len(mytablef.Fields("DUA"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("DUA"), nuevoDato) 'DUA 92
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("cantdev") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("cantdev"))
                myDato = mytablef.Fields("cantdev")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'cantdev 93
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("servicioco") = 0 And mytablef.Fields("T6") = "?" Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("servicioco"))
                myDato = mytablef.Fields("servicioco")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'servicioco 94
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("serviciopo") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("serviciopo"))
                myDato = mytablef.Fields("serviciopo")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'serviciopo 95
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("destopo") = 0 Then
                myREG = myREG & Space$(7) & 0
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("destopo"))
                myDato = mytablef.Fields("destopo")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'destopo 96
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("fechaborra")) = "vacio" Then
                myREG = myREG & Space$(10)
                myREG = myREG & "&"
            Else
                hastaCuanto = 10 - Len(mytablef.Fields("fechaborra"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("fechaborra"), nuevoDato) 'fechaborra 97
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("horaborra")) = "vacio" Then
                myREG = myREG & Space$(10)
                myREG = myREG & "&"
            Else
                hastaCuanto = 10 - Len(mytablef.Fields("horaborra"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("horaborra"), nuevoDato) 'horaborra 98
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("administrador")) = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("administrador"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("administrador"), nuevoDato) 'administrador 99
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("detraccion") = 0 Then
                myREG = myREG & Space$(7) & 0
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("detraccion"))
                myDato = mytablef.Fields("detraccion")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'detraccion 100
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            Print #Filelibero1, myREG
            Close #Filelibero1
            mytablef.MoveNext

            'aqui hace el progress bar
            Frm_backup.ProgressBar1.Value = ((conta / conta_record) * 100)
            Frm_backup.lblElaborandoBackup = "Elaborando backup al.." & ((conta / conta_record) * 100) & "%"
        Loop
        Close #Filelibero1

    End If

    mytablef.Close

End Sub

Public Function llenar_datos(hastaCuanto As Integer, _
                             myDato As String, _
                             nuevoDato As String)

    For I = 1 To hastaCuanto
        myDato = Space$(1) + myDato
    Next
    nuevoDato = myDato

End Function

Public Function read_cantidad_file_enviado(input_file As String, cantidad As Long, fnum)

    Dim input_record As String

    On Error GoTo read_cantidad_file_enviado

    fnum = FreeFile

    cantidad = 0

    Open input_file For Input As #fnum

    Do Until EOF(fnum)
        Line Input #fnum, input_record
        cantidad = cantidad + 1
    Loop
 
    Close #fnum
    cantidad = cantidad - 1

read_cantidad_file_enviado:
    Exit Function

End Function

Public Function read_save_detalle(input_file As String, cantidad As Long)

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
        mysql = "INSERT INTO detalle " & Chr$(10)
        mysql = mysql & "(tipo," & Chr$(10)
        mysql = mysql & "serie," & Chr$(10)
        mysql = mysql & "numero," & Chr$(10)
        mysql = mysql & "tipoclie," & Chr$(10)
        mysql = mysql & "codigo," & Chr$(10)
        mysql = mysql & "ACU," & Chr$(10)
        mysql = mysql & "acu1," & Chr$(10)
        mysql = mysql & "fecha," & Chr$(10)
        mysql = mysql & "moneda, " & Chr$(10)
        mysql = mysql & "PRODUCTO," & Chr$(10)
        mysql = mysql & "DESCRIPCIO," & Chr$(10)
        mysql = mysql & "unidad ," & Chr$(10)
        mysql = mysql & "factor ," & Chr$(10)
        mysql = mysql & "cantidad , " & Chr$(10)
        mysql = mysql & "PRECIO," & Chr$(10)
        mysql = mysql & "IGV," & Chr$(10)
        mysql = mysql & "neto ," & Chr$(10)
        mysql = mysql & "descuento ," & Chr$(10)
        mysql = mysql & "subtotal ," & Chr$(10)
        mysql = mysql & "impuesto , " & Chr$(10)
        mysql = mysql & "TOTAL," & Chr$(10)
        mysql = mysql & "ESTADO," & Chr$(10)
        mysql = mysql & "usuario ," & Chr$(10)
        mysql = mysql & "FECHACREA ," & Chr$(10)
        mysql = mysql & "hora ," & Chr$(10)
        mysql = mysql & "vendedor ," & Chr$(10)
        mysql = mysql & "bodega , " & Chr$(10)
        mysql = mysql & "BODEGAF," & Chr$(10)
        mysql = mysql & "DESLIPO," & Chr$(10)
        mysql = mysql & "flage ," & Chr$(10)
        mysql = mysql & "linea ," & Chr$(10)
        mysql = mysql & "t1 ," & Chr$(10)
        mysql = mysql & "t2 ," & Chr$(10)
        mysql = mysql & "t3 ," & Chr$(10)
        mysql = mysql & "t4 ," & Chr$(10)
        mysql = mysql & "t5 ," & Chr$(10)
        mysql = mysql & "t6 ," & Chr$(10)
        mysql = mysql & "t7 ," & Chr$(10)
        mysql = mysql & "t8 ," & Chr$(10)
        mysql = mysql & "t9 , " & Chr$(10)
        mysql = mysql & "T10," & Chr$(10)
        mysql = mysql & "T11," & Chr$(10)
        mysql = mysql & "t12 ," & Chr$(10)
        mysql = mysql & "t13 ," & Chr$(10)
        mysql = mysql & "t14 ," & Chr$(10)
        mysql = mysql & "t15 ," & Chr$(10)
        mysql = mysql & "t16 , " & Chr$(10)
        mysql = mysql & "L1," & Chr$(10)
        mysql = mysql & "L2," & Chr$(10)
        mysql = mysql & "L3 ," & Chr$(10)
        mysql = mysql & "L4 , " & Chr$(10)
        mysql = mysql & "LOCAL," & Chr$(10)
        mysql = mysql & "PROVEEDORP," & Chr$(10)
        mysql = mysql & "observa1 ," & Chr$(10)
        mysql = mysql & "observa2 ," & Chr$(10)
        mysql = mysql & "observa3 ," & Chr$(10)
        mysql = mysql & "observa4 ," & Chr$(10)
        mysql = mysql & "zona , " & Chr$(10)
        mysql = mysql & "ISC," & Chr$(10)
        mysql = mysql & "TAX," & Chr$(10)
        mysql = mysql & "vtaneta ," & Chr$(10)
        mysql = mysql & "tcosto ," & Chr$(10)
        mysql = mysql & "ganancia ," & Chr$(10)
        mysql = mysql & "comision ," & Chr$(10)
        mysql = mysql & "cajero , " & Chr$(10)
        mysql = mysql & "CAJA," & Chr$(10)
        mysql = mysql & "turno," & Chr$(10)
        mysql = mysql & "servicio, " & Chr$(10)
        mysql = mysql & "comanda," & Chr$(10)
        mysql = mysql & "mesa ," & Chr$(10)
        mysql = mysql & "salon ," & Chr$(10)
        mysql = mysql & "mesero ," & Chr$(10)
        mysql = mysql & "sentido ," & Chr$(10)
        mysql = mysql & "ccosto ," & Chr$(10)
        mysql = mysql & "familia ," & Chr$(10)
        mysql = mysql & "subfamilia , " & Chr$(10)
        mysql = mysql & "MARCA," & Chr$(10)
        mysql = mysql & "PERCEPCION," & Chr$(10)
        mysql = mysql & "TPERCEPCIO ," & Chr$(10)
        mysql = mysql & "flete ," & Chr$(10)
        mysql = mysql & "localf ," & Chr$(10)
        mysql = mysql & "ivap ," & Chr$(10)
        mysql = mysql & "tivap ," & Chr$(10)
        mysql = mysql & "NROPRECIO ," & Chr$(10)
        mysql = mysql & "tisc ," & Chr$(10)
        mysql = mysql & "PLACA ," & Chr$(10)
        mysql = mysql & "xneto ," & Chr$(10)
        mysql = mysql & "tdetra ," & Chr$(10)
        mysql = mysql & "DENUMERO , " & Chr$(10)
        mysql = mysql & "CATEGORIA," & Chr$(10)
        mysql = mysql & "ADUANA," & Chr$(10)
        mysql = mysql & "DUA ," & Chr$(10)
        mysql = mysql & "cantdev ," & Chr$(10)
        mysql = mysql & "servicioco ," & Chr$(10)
        mysql = mysql & "serviciopo ," & Chr$(10)
        mysql = mysql & "destopo ," & Chr$(10)
        mysql = mysql & "fechaborra , " & Chr$(10)
        mysql = mysql & "horaborra," & Chr$(10)
        mysql = mysql & "administrador," & Chr$(10)
        mysql = mysql & "detraccion)" & Chr$(10)
 
        mysql = mysql & " VALUES ('" & Mid(input_record, 1, 3) & "'," & Chr$(10) 'tipo 1
        mysql = mysql & " '" & Mid(input_record, 5, 4) & "'," & Chr$(10) 'SERIE 2
        mysql = mysql & " '" & Trim(Mid(input_record, 10, 11)) & "'," & Chr$(10) 'NUMERO 3
        mysql = mysql & " '" & Trim(Mid(input_record, 22, 1)) & "'," & Chr$(10) 'tipoclie 4
    
        mysql = mysql & " '" & Trim(Mid(input_record, 24, 11)) & "'," & Chr$(10) 'codigo 5
    
        If (Mid(input_record, 36, 1)) = Space$(1) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & " '" & Trim$(Mid(input_record, 36, 1)) & "'," & Chr$(10) 'ACU 6

        End If
        
        If (Mid(input_record, 38, 1)) = Space$(1) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & " '" & Trim$(Mid(input_record, 38, 1)) & "'," & Chr$(10) 'ACU1 7

        End If
    
        If (Mid(input_record, 40, 10)) = Space$(10) Then
            mysql = mysql & Space(10) & Chr$(10)
        Else
            mysql = mysql & " '" & Trim$(Mid(input_record, 40, 10)) & "'," & Chr$(10) 'FECHA 8

        End If

        mysql = mysql & " '" & Trim$(Mid(input_record, 51, 1)) & "'," & Chr$(10) 'MONEDA 9
        mysql = mysql & " '" & Trim$(Mid(input_record, 53, 15)) & "'," & Chr$(10) 'PRODUCTO 10
    
        If Mid(input_record, 69, 120) = Space$(120) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            myDato = Trim(Mid(input_record, 69, 120))
            myCuenta = Len(Trim(Mid(input_record, 69, 120)))
            newDESCRIPCIO = Replace$(myDato, "'", Chr(34))
            'mysql = mysql & " '" & Trim(Mid(input_record, 69, 120)) & "'," & Chr$(10) 'DESCRIPCIO 11
            mysql = mysql & " '" & newDESCRIPCIO & "'," & Chr$(10) 'DESCRIPCIO 11

        End If
    
        mysql = mysql & " '" & Trim(Mid(input_record, 190, 6)) & "'," & Chr$(10) 'UNIDAD 12

        'aqui solo numero
        If Trim(Mid(input_record, 197, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim(Mid(input_record, 197, 8)) & "," & Chr$(10) 'FACTOR" 13

        End If
    
        If (Mid(input_record, 206, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 206, 8)) & "," & Chr$(10) 'cantidad 14

        End If
    
        If (Mid(input_record, 215, 8)) = Space$(8) And Trim((Mid(input_record, 215, 8))) = "" Then
            mysql = mysql & 0 & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 215, 8)) & "," & Chr$(10) 'PRECIO 15

        End If
    
        If (Mid(input_record, 224, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 224, 8)) & "," & Chr$(10) 'IGV 16

        End If
    
        If (Mid(input_record, 233, 10)) = Space$(10) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 233, 10)) & "," & Chr$(10) 'NETO 17

        End If
    
        If (Mid(input_record, 244, 10)) = Space$(10) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 244, 10)) & "," & Chr$(10) 'DESCUENTO 18

        End If
    
        If (Mid(input_record, 255, 12)) = Space$(12) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 255, 12)) & "," & Chr$(10) 'SUBTOTAL 19

        End If
    
        If (Mid(input_record, 268, 12)) = Trim(0) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 268, 12)) & "," & Chr$(10) 'IMPUESTO 20

        End If
    
        If Mid(input_record, 281, 12) = Space$(12) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 281, 12)) & "," & Chr$(10) 'TOTAL 21

        End If
    
        If Mid(input_record, 294, 1) = Space$(1) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 294, 1)) & "'," & Chr$(10) 'ESTADO 22

        End If

        If Mid(input_record, 296, 12) = Space$(12) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim(Mid(input_record, 296, 12)) & "'," & Chr$(10) 'USUARIO 23

        End If
    
        If Mid(input_record, 309, 10) = Space$(10) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 309, 10)) & "'," & Chr$(10) 'FECHACREA 24

        End If
    
        If Mid(input_record, 320, 8) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 320, 8)) & "'," & Chr$(10) 'HORA 25

        End If

        If (Mid(input_record, 329, 11)) = Space$(11) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 329, 11)) & "'," & Chr$(10) 'VENDEDOR 26

        End If
    
        If Mid(input_record, 341, 2) = Space$(2) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 341, 2)) & "'," & Chr$(10) 'BODEGA 27

        End If
    
        If Mid(input_record, 344, 2) = Space$(2) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 344, 2)) & "'," & Chr$(10) 'BODEGAF 28

        End If
    
        If Mid(input_record, 347, 10) = Trim$(0) Or Mid(input_record, 347, 10) = Space$(10) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 347, 10)) & "'," & Chr$(10) 'DESLIPO 29

        End If
    
        If Mid(input_record, 358, 1) = Space$(1) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 358, 1)) & "'," & Chr$(10) 'FLAGE 30

        End If
    
        If Mid(input_record, 360, 6) = Space$(6) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 360, 6)) & "'," & Chr$(10) 'LINEA 31

        End If

        '    aqui numeros
        If (Mid(input_record, 367, 8)) = Space(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 367, 8)) & "," & Chr$(10) 'T1 32

        End If

        If (Mid(input_record, 376, 8)) = Space(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 376, 8)) & "," & Chr$(10) 'T2 33

        End If

        If (Mid(input_record, 385, 8)) = Space(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 385, 8)) & "," & Chr$(10) 'T3 34

        End If

        If (Mid(input_record, 394, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 394, 8)) & "," & Chr$(10) 'T4 35

        End If

        If (Mid(input_record, 403, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 403, 8)) & "," & Chr$(10) 'T5 36

        End If

        If (Mid(input_record, 412, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 412, 8)) & "," & Chr$(10) 'T6 37

        End If

        If (Mid(input_record, 421, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 421, 8)) & "," & Chr$(10) 'T7 38

        End If

        If (Mid(input_record, 430, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 430, 8)) & "," & Chr$(10) 'T8 39

        End If

        If (Mid(input_record, 439, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 439, 8)) & "," & Chr$(10) 'T9 40

        End If

        If (Mid(input_record, 448, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 448, 8)) & "," & Chr$(10) 'T10 41

        End If

        If (Mid(input_record, 457, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 457, 8)) & "," & Chr$(10) 'T11 42

        End If

        If (Mid(input_record, 466, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 466, 8)) & "," & Chr$(10) 'T12 43

        End If
    
        If (Mid(input_record, 475, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 475, 8)) & "," & Chr$(10) 'T13 44

        End If

        If (Mid(input_record, 484, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 484, 8)) & "," & Chr$(10) 'T14 45

        End If

        If (Mid(input_record, 493, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 493, 8)) & "," & Chr$(10) 'T15 46

        End If

        If (Mid(input_record, 502, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 502, 8)) & "," & Chr$(10) 'T16 47

        End If

        If (Mid(input_record, 511, 1)) = Space$(1) Or (Mid(input_record, 511, 1)) = "?" Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 511, 1)) & "'," & Chr$(10) 'L1 48

        End If

        If (Mid(input_record, 513, 1)) = Space$(1) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 513, 1)) & "'," & Chr$(10) 'L2 49

        End If
    
        If (Mid(input_record, 515, 1)) = Space$(1) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 515, 1)) & "'," & Chr$(10) 'L3 50

        End If

        If (Mid(input_record, 517, 1)) = Space$(1) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 517, 1)) & "'," & Chr$(10) 'L4 51

        End If

        mysql = mysql & "'" & Trim$((Mid(input_record, 519, 6))) & "'," & Chr$(10) 'LOCAL 52
        mysql = mysql & "'" & Trim$(Mid(input_record, 526, 11)) & "'," & Chr$(10) 'PROVEEDORP 53

        If (Mid(input_record, 538, 150)) = Space$(150) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 538, 150)) & "'," & Chr$(10) 'OBSERVA1 54

        End If

        If (Mid(input_record, 689, 150)) = Space$(150) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 689, 150)) & "'," & Chr$(10) 'OBSERVA2 55

        End If
    
        If (Mid(input_record, 840, 150)) = Space$(150) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 840, 150)) & "'," & Chr$(10) 'OBSERVA3 56

        End If

        If (Mid(input_record, 991, 150)) = Space$(150) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 991, 150)) & "'," & Chr$(10) 'OBSERVA4 57

        End If

        If (Mid(input_record, 1142, 6)) = Space$(6) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1142, 6)) & "'," & Chr$(10) 'zona 58

        End If

        'aqui numeros
    
        If (Mid(input_record, 1149, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 1149, 8)) & "," & Chr$(10) 'ISC 59

        End If
    
        If (Mid(input_record, 1158, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 1158, 8)) & "," & Chr$(10) 'TAX 60

        End If
    
        If (Mid(input_record, 1167, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 1167, 8)) & "," & Chr$(10) 'VTANETA 61

        End If
    
        If (Mid(input_record, 1176, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 1176, 8)) & "," & Chr$(10) 'TCOSTO 62

        End If

        If (Mid(input_record, 1185, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 1185, 8)) & "," & Chr$(10) 'GANANCIA 63

        End If
    
        If (Mid(input_record, 1194, 10)) = Space(10) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 1194, 10)) & "," & Chr$(10) 'COMISION 64

        End If

        If (Mid(input_record, 1205, 11)) = Space$(11) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1205, 11)) & "'," & Chr$(10) 'CAJERO 65

        End If
    
        If Mid(input_record, 1217, 3) = Space$(3) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1217, 3)) & "'," & Chr$(10) 'CAJA 66

        End If
    
        If (Mid(input_record, 1221, 1)) = Space$(1) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1221, 1)) & "'," & Chr$(10) 'TURNO 67

        End If

        If (Mid(input_record, 1223, 1)) = Space$(1) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1223, 1)) & "'," & Chr$(10) 'SERVICIO 68

        End If

        If (Mid(input_record, 1225, 11)) = Space$(11) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1225, 11)) & "'," & Chr$(10) 'COMANDA 69

        End If
    
        If (Mid(input_record, 1237, 3)) = Space$(3) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1237, 3)) & "'," & Chr$(10) 'MESA 70

        End If

        If (Mid(input_record, 1241, 3)) = Space$(3) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1241, 3)) & "'," & Chr$(10) 'SALON 71

        End If

        If (Mid(input_record, 1245, 11)) = Space$(11) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1245, 11)) & "'," & Chr$(10) 'MESERO 72

        End If
    
        If Mid(input_record, 1257, 1) = Space$(1) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1257, 1)) & "'," & Chr$(10) 'SENTIDO 73

        End If
    
        If (Mid(input_record, 1259, 6)) = Space(6) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1259, 6)) & "'," & Chr$(10) 'CCOSTO 74

        End If

        If (Mid(input_record, 1266, 6)) = Space$(6) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1266, 6)) & "'," & Chr$(10) 'FAMILIA 75

        End If

        If (Mid(input_record, 1273, 6)) = Space$(6) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1273, 6)) & "'," & Chr$(10) 'SUBFAMILIA 76

        End If

        If (Mid(input_record, 1280, 6)) = Space$(6) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1280, 6)) & "'," & Chr$(10) 'MARCA 77

        End If

        If (Mid(input_record, 1287, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 1287, 8)) & "," & Chr$(10) 'PERCEPCION 78

        End If

        If (Mid(input_record, 1296, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 1296, 8)) & "," & Chr$(10) 'TPERCEPCIO 79

        End If

        If (Mid(input_record, 1305, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 1305, 8)) & "," & Chr$(10) 'FLETE 80

        End If

        mysql = mysql & "'" & Trim$(Mid(input_record, 1314, 3)) & "'," & Chr$(10) 'LOCALF 81

        If (Mid(input_record, 1318, 3)) = Space$(3) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 1318, 3)) & "," & Chr$(10) 'IVAP 82

        End If

        If (Mid(input_record, 1327, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 1327, 8)) & "," & Chr$(10) 'TIVAP 83

        End If

        If Mid(input_record, 1336, 1) = Space$(1) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1336, 1)) & "'," & Chr$(10) 'NROPRECIO 84

        End If
    
        If (Mid(input_record, 1338, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 1338, 8)) & "," & Chr$(10) 'TCIS 85

        End If

        If (Mid(input_record, 1347, 15)) = Space(15) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1347, 15)) & "'," & Chr$(10) 'PLACA 86

        End If

        If (Mid(input_record, 1363, 8)) = Space(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 1363, 8)) & "," & Chr$(10) 'XNETO 87

        End If

        If (Mid(input_record, 1372, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 1372, 8)) & "," & Chr$(10) 'TDETRA 88

        End If

        If (Mid(input_record, 1381, 11)) = Space$(11) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1381, 11)) & "'," & Chr$(10) 'DENUMERO 89

        End If

        If (Mid(input_record, 1393, 6)) = Space$(6) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1393, 6)) & "'," & Chr$(10) 'CATEGORIA 90

        End If

        If (Mid(input_record, 1400, 11)) = Space(11) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1400, 11)) & "'," & Chr$(10) 'ADUANA 91

        End If

        If (Mid(input_record, 1412, 10)) = Space$(10) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1412, 10)) & "'," & Chr$(10) 'DUA 92

        End If

        If (Mid(input_record, 1423, 8)) = Space(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 1423, 8)) & "," & Chr$(10) 'cantdev 93

        End If

        If (Mid(input_record, 1432, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 1432, 8)) & "," & Chr$(10) 'servicioco 94

        End If

        If (Mid(input_record, 1441, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 1441, 8)) & "," & Chr$(10) 'serviciopo 95

        End If

        If (Mid(input_record, 1450, 8)) = Space$(8) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 1450, 8)) & "," & Chr$(10) 'destopo 96

        End If

        If (Mid(input_record, 1459, 10)) = Space(10) Or Trim(Mid(input_record, 1459, 10)) = 0 Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1459, 10)) & "'," & Chr$(10) 'fechaborra 97

        End If

        If (Mid(input_record, 1470, 10)) = Space(10) Or (Mid(input_record, 1470, 10)) = 0 Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1470, 10)) & "'," & Chr$(10) 'horaborra 98

        End If

        If (Mid(input_record, 1481, 11)) = Space(11) Then
            mysql = mysql & "NULL" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1481, 11)) & "'," & Chr$(10) 'administrador 99

        End If

        If (Mid(input_record, 1493, 8)) = Space(8) Then
            mysql = mysql & "NULL" & ")" & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 1493, 8)) & ")" & Chr$(10) 'detraccion '100

        End If
        
        'mytablef.Open mysql, cn, adOpenStatic, adLockOptimistic
        cn.Execute (mysql)

        If cantidad < my_conta Then
            Exit Do

        End If

        Frm_backup.ProgressBar1.Value = ((my_conta / cantidad) * 100)
        Frm_backup.lblElaborandoBackup.Caption = "Actualizando al.." & ((my_conta / cantidad) * 100) & "%"
        'mytablef.MoveLast

    Loop
    Close #fnum

End Function

'inicio 02/05/2017 pll
'inicio 25/04/2017 pll
Public Sub bkp_factura(fechai As String, fechaf As String)

    Dim mysql       As String

    Dim mytablef    As New ADODB.Recordset

    Dim hastaCuanto As Integer

    Dim myDato      As String

    Dim nuevoDato   As String

    Dim my_Lnumero  As Integer

    Dim conta       As Integer

    Dim myREG       As String

    FileName = "C:\EmpaquetaVi\" & "bkpfactura" & ".txt"

    mysql = "SELECT isnull(tipo,'vac') as tipo," & Chr$(10)
    mysql = mysql & "isnull(serie,'vaci') as serie , " & Chr$(10)
    mysql = mysql & "isnull(numero,'vacio') as numero , " & Chr$(10)
    mysql = mysql & "isnull(tipoclie,'v') as tipoclie, " & Chr$(10)
    mysql = mysql & "isnull(codigo,'vacio') as codigo, " & Chr$(10)
    mysql = mysql & "isnull(partida,'vacio') as partida, " & Chr$(10)
    mysql = mysql & "isnull(destino,'vacio') as destino," & Chr$(10)
    mysql = mysql & "isnull(fecha,0) as fecha," & Chr$(10)
    mysql = mysql & "isnull(fechae,0) as fechae," & Chr$(10)
    mysql = mysql & "isnull(moneda,'?') as moneda," & Chr$(10)
    mysql = mysql & "isnull(vendedor,'vacio') as vendedor," & Chr$(10)
    mysql = mysql & "isnull(transporte,'vacio') as transporte," & Chr$(10)
    mysql = mysql & "isnull(fpago,'vac') as fpago," & Chr$(10)
    mysql = mysql & "isnull(paridad,0) as paridad," & Chr$(10)
    mysql = mysql & "isnull(dias,0) as dias," & Chr$(10)
    mysql = mysql & "isnull(bodega,'va') as bodega," & Chr$(10)
    mysql = mysql & "isnull(bodegaf,'va') as bodegaf," & Chr$(10)
    mysql = mysql & "isnull(observa,'vacio') as observa," & Chr$(10)
    mysql = mysql & "isnull(estado,'v') as estado," & Chr$(10)
    mysql = mysql & "isnull(acu,'?') as acu," & Chr$(10)
    mysql = mysql & "isnull(acu1,'?') as  acu1," & Chr$(10)
    mysql = mysql & "isnull(usuario,'vacio') as usuario," & Chr$(10)
    mysql = mysql & "isnull(fechacrea,0) as fechacrea," & Chr$(10)
    mysql = mysql & "isnull(hora,0) as hora," & Chr$(10)
    mysql = mysql & "isnull(nombre,'vacio') as nombre," & Chr$(10)
    mysql = mysql & "isnull(total,0)as total," & Chr$(10)
    mysql = mysql & "isnull(descuento,0) as descuento," & Chr$(10)
    mysql = mysql & "isnull(neto,0) as neto," & Chr$(10)
    mysql = mysql & "isnull(impuesto,0) as impuesto," & Chr$(10)
    mysql = mysql & "isnull(subtotal,0) as subtotal," & Chr$(10)
    mysql = mysql & "isnull(flage,'?') as flage," & Chr$(10)
    mysql = mysql & "isnull(tipo1,'vac') as tipo1," & Chr$(10)
    mysql = mysql & "isnull(serie1,'vaci') as serie1," & Chr$(10)
    mysql = mysql & "isnull(numero1,'vacio') as numero1," & Chr$(10)
    mysql = mysql & "isnull(serie2,'vaci') as serie2," & Chr$(10)
    mysql = mysql & "isnull(numero2,'vacio') as numero2," & Chr$(10)
    mysql = mysql & "isnull(serie3,'vaci') as serie3," & Chr$(10)
    mysql = mysql & "isnull(numero3,'vacio') as numero3," & Chr$(10)
    mysql = mysql & "isnull(serie4,'vaci') as serie4," & Chr$(10)
    mysql = mysql & "isnull(numero4,'vacio') as numero4," & Chr$(10)
    mysql = mysql & "isnull(serie5,'vaci') as serie5," & Chr$(10)
    mysql = mysql & "isnull(numero5,'vacio') as numero5," & Chr$(10)
    mysql = mysql & "isnull(serie6,'vaci') as serie6," & Chr$(10)
    mysql = mysql & "isnull(numero6,'vacio') as numero6," & Chr$(10)
    mysql = mysql & "isnull(serie7,'vaci') as serie7," & Chr$(10)
    mysql = mysql & "isnull(numero7,'vacio') as numero7," & Chr$(10)
    mysql = mysql & "isnull(serie8,'vacio') as serie8," & Chr$(10)
    mysql = mysql & "isnull(numero8,'vacio') as numero8," & Chr$(10)
    mysql = mysql & "isnull(numero8,'vacio') as nop," & Chr$(10)
    mysql = mysql & "isnull(local,'vacio') as local," & Chr$(10)
    mysql = mysql & "isnull(c1,0) as c1," & Chr$(10)
    mysql = mysql & "isnull(c2,0) as c2," & Chr$(10)
    mysql = mysql & "isnull(c3,0) as c3," & Chr$(10)
    mysql = mysql & "isnull(c4,0) as c4," & Chr$(10)
    mysql = mysql & "isnull(c5,0) as c5," & Chr$(10)
    mysql = mysql & "isnull(c6,0) as c6," & Chr$(10)
    mysql = mysql & "isnull(c7,0) as c7," & Chr$(10)
    mysql = mysql & "isnull(c8,0) as c8," & Chr$(10)
    mysql = mysql & "isnull(c9,0) as c9," & Chr$(10)
    mysql = mysql & "isnull(zona,'vacio') as zona," & Chr$(10)
    mysql = mysql & "isnull(retipo1,'vacio') as retipo1," & Chr$(10)
    mysql = mysql & "isnull(renumero1,'vacio') as renumero1," & Chr$(10)
    mysql = mysql & "isnull(renumero2,'vacio') as renumero2," & Chr$(10)
    mysql = mysql & "isnull(renumero3,'vacio') as renumero3," & Chr$(10)
    mysql = mysql & "isnull(retotal,0) as retotal," & Chr$(10)
    mysql = mysql & "isnull(retotal1,0) as retotal1," & Chr$(10)
    mysql = mysql & "isnull(retotal2,0) as retotal2," & Chr$(10)
    mysql = mysql & "isnull(retotal3,0) as retotal3," & Chr$(10)
    mysql = mysql & "isnull(acuenta,0) as acuenta," & Chr$(10)
    mysql = mysql & "isnull(nro_items,0) as nro_items," & Chr$(10)
    mysql = mysql & "isnull(acuenta,0) as acuenta," & Chr$(10)
    mysql = mysql & "isnull(adetotal,0) as adetotal," & Chr$(10)
    mysql = mysql & "isnull(yausado,'v') as yausado," & Chr$(10)
    mysql = mysql & "isnull(caja,'vac') as caja," & Chr$(10)
    mysql = mysql & "isnull(turno,'?') as turno," & Chr$(10)
    mysql = mysql & "isnull(servicio,'?') as servicio," & Chr$(10)
    mysql = mysql & "isnull(comanda,'vacio') as comanda," & Chr$(10)
    mysql = mysql & "isnull(mesa,'vacio') as mesa," & Chr$(10)
    mysql = mysql & "isnull(salon,'vacio') as salon," & Chr$(10)
    mysql = mysql & "isnull(mesero,'vacio') as mesero," & Chr$(10)
    mysql = mysql & "isnull(telefono,'vacio') as telefono," & Chr$(10)
    mysql = mysql & "isnull(ruc,'vacio') as ruc," & Chr$(10)
    mysql = mysql & "isnull(montopagar,0) as montopagar," & Chr$(10)
    mysql = mysql & "isnull(tdocdeli,'vacio') as tdocdeli," & Chr$(10)
    mysql = mysql & "isnull(gravado,0) as gravado," & Chr$(10)
    mysql = mysql & "isnull(fechasunat,0) as fechasunat," & Chr$(10)
    mysql = mysql & "isnull(flag_deli,'?') as flag_deli," & Chr$(10)
    mysql = mysql & "isnull(redondeo,0) as redondeo," & Chr$(10)
    mysql = mysql & "isnull(percepcion,0)as  percepcion," & Chr$(10)
    mysql = mysql & "isnull(flag_deli,0) as tflete," & Chr$(10)
    mysql = mysql & "isnull(localf,'vacio') as localf," & Chr$(10)
    mysql = mysql & "isnull(tivap,0) as tivap," & Chr$(10)
    mysql = mysql & "isnull(tisc,0) as tisc," & Chr$(10)
    mysql = mysql & "isnull(telefono,'vacio') as placa," & Chr$(10)
    mysql = mysql & "isnull(xneto,0) as xneto," & Chr$(10)
    mysql = mysql & "isnull(tdetra,0) as tdetra," & Chr$(10)
    mysql = mysql & "isnull(sentido,'?') as sentido," & Chr$(10)
    mysql = mysql & "isnull(denumero,'vacio') as denumero," & Chr$(10)
    mysql = mysql & "isnull(denumero,'vacio') as dflag," & Chr$(10)
    mysql = mysql & "isnull(aduana,'vacio') as aduana," & Chr$(10)
    mysql = mysql & "isnull(aduana,'vacio') as dua," & Chr$(10)
    mysql = mysql & "isnull(importacio,'vacio') as importacio," & Chr$(10)
    mysql = mysql & "isnull(tipoimp,'vacio') as tipoimp," & Chr$(10)
    mysql = mysql & "isnull(serieimp,'vacio') as serieimp," & Chr$(10)
    mysql = mysql & "isnull(numeroimp,'vacio') as numeroimp," & Chr$(10)
    mysql = mysql & "isnull(numeroimp,'vacio') as gasto," & Chr$(10)
    mysql = mysql & "isnull(servicioco,0) as servicioco," & Chr$(10)
    mysql = mysql & "isnull(clasesunat,'vacio') as clasesunat," & Chr$(10)
    mysql = mysql & "isnull(destopo,0) as destopo," & Chr$(10)
    mysql = mysql & "isnull(horae,'vacio') as horae," & Chr$(10)
    mysql = mysql & "isnull(vendedor1,'vacio') as vendedor1," & Chr$(10)
    mysql = mysql & "isnull(vendedor2,'vacio') as vendedor2," & Chr$(10)
    mysql = mysql & "isnull(vendedor3,'vacio') as vendedor3," & Chr$(10)
    mysql = mysql & "isnull(vendedor4,'vacio') as vendedor4," & Chr$(10)
    mysql = mysql & "isnull(codigo1,'vacio') as codigo1," & Chr$(10)
    mysql = mysql & "isnull(personas,0) as personas" & Chr$(10)
    mysql = mysql & "from factura where fecha>='" & Format(fechai, "YYYYMMDD") & "'" & Chr$(10)
    mysql = mysql & "and fecha<='" & Format(fechaf, "YYYYMMDD") & "' " & Chr$(10)
    mysql = mysql & "order by fecha" & Chr$(10)

    mytablef.Open mysql, cn, adOpenStatic, adLockOptimistic
  
    If mytablef.RecordCount > 0 Then
        Do

            If mytablef.EOF Then Exit Do
            'aqui es para abrir el file txt y llenarlo
            Filelibero1 = FreeFile
            Open FileName For Append As #Filelibero1
            conta = conta + 1
      
            conta_record = mytablef.RecordCount
            myREG = ""
      
            If mytablef.Fields("tipo") = "vac" Then
                myREG = Space$(3)
                myREG = myREG & "&"
            Else
                hastaCuanto = 3 - Trim(Len(mytablef.Fields("tipo")))
                Call llenar_datos(hastaCuanto, mytablef.Fields("tipo"), nuevoDato) 'tipo 1
                myREG = myREG & nuevoDato
                myREG = myREG & "&" 'gion separador

            End If
      
            If mytablef.Fields("SERIE") = "vaci" Then
                myREG = Space$(4)
                myREG = myREG & "&"
            Else
                hastaCuanto = 4 - Trim(Len(mytablef.Fields("SERIE")))
                Call llenar_datos(hastaCuanto, mytablef.Fields("SERIE"), nuevoDato) 'SERIE 2
                myREG = myREG & nuevoDato 'SERIE 2
                myREG = myREG & "&" 'gion separador

            End If
      
            If mytablef.Fields("NUMERO") = "vacio" Then
                myREG = Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("NUMERO"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("NUMERO"), nuevoDato) 'NUMERO 3
                myREG = myREG & nuevoDato
                myREG = myREG & "&" 'gion separador

            End If
      
            If mytablef.Fields("tipoclie") = "v" Then
                myREG = Space$(1)
                myREG = myREG & "&"
            Else
                myREG = myREG & mytablef.Fields("tipoclie") 'tipoclie 4
                myREG = myREG & "&" 'gion separador

            End If
      
            If mytablef.Fields("codigo") = "vacio" Then
                myREG = Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("codigo"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("codigo"), nuevoDato) 'codigo 5
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("partida") = "vacio" Then
                myREG = Space$(60)
                myREG = myREG & "&"
            Else
                hastaCuanto = 60 - Len(mytablef.Fields("partida"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("partida"), nuevoDato) 'partida 6
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("destino") = "vacio" Then
                myREG = Space$(60)
                myREG = myREG & "&"
            Else
                hastaCuanto = 60 - Len(mytablef.Fields("destino"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("destino"), nuevoDato) 'destino 7
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("FECHA") = 0 Then
                myREG = Space$(10)
                myREG = myREG & "&"
            Else
                myREG = myREG & mytablef.Fields("FECHA") 'FECHA 8
                myREG = myREG & "&" 'gion separador

            End If
      
            If mytablef.Fields("fechae") = 0 Then
                myREG = Space$(10)
                myREG = myREG & "&"
            Else
                myREG = myREG & mytablef.Fields("fechae") 'fechae 9
                myREG = myREG & "&" 'gion separador

            End If
      
            If mytablef.Fields("MONEDA") = "?" Then
                myREG = Space$(1)
                myREG = myREG & "&"
            Else
                myREG = myREG & mytablef.Fields("MONEDA") 'MONEDA 10
                myREG = myREG & "&" 'gion separador

            End If
      
            If Trim(mytablef.Fields("VENDEDOR")) = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("VENDEDOR"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("VENDEDOR"), nuevoDato) 'VENDEDOR 11
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("transporte")) = "" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("transporte"))
                Call llenar_datos(hastaCuanto, Len(mytablef.Fields("transporte")), nuevoDato) 'transporte 12
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("FPAGO")) = "vac" Then
                myREG = myREG & Space$(3)
                myREG = myREG & "&"
            Else
                hastaCuanto = 3 - Len(mytablef.Fields("FPAGO"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("FPAGO"), nuevoDato) 'FPAGO 13
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("PARIDAD") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("PARIDAD"))
                myDato = mytablef.Fields("PARIDAD")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'PARIDAD 14
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("DIAS") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("DIAS"))
                myDato = mytablef.Fields("DIAS")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'DIAS 15
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("BODEGA") = "va" Then
                myREG = myREG & Space$(2)
                myREG = myREG & "&"
            Else
                myREG = myREG & mytablef.Fields("BODEGA")                'BODEGA 16
                myREG = myREG & "&" 'gion separador

            End If
      
            If Trim(mytablef.Fields("bodegaf")) = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 2 - Len(mytablef.Fields("bodegaf"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("bodegaf"), nuevoDato) 'bodegaf 17
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("observa")) = "vacio" Then
                myREG = myREG & Space$(60)
                myREG = myREG & "&"
            Else
                hastaCuanto = 60 - Len(mytablef.Fields("observa"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("observa"), nuevoDato) 'observa 18
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("estado") = "v" Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                myREG = myREG & mytablef.Fields("estado") 'estado 19
                myREG = myREG & "&" 'gion separador

            End If
      
            If mytablef.Fields("acu") = "?" Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                myREG = myREG & mytablef.Fields("acu") 'acu 20
                myREG = myREG & "&" 'gion separador

            End If
      
            If mytablef.Fields("ACU1") = "" Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                myREG = myREG & mytablef.Fields("acu1")   'acu1 21
                myREG = myREG & "&" 'gion separador

            End If

            If Trim(mytablef.Fields("usuario")) = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("usuario"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("usuario"), nuevoDato) 'usuario 22
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("fechacrea") = 0 Then
                myREG = myREG & Space$(10)
                myREG = myREG & "&"
            Else
                myREG = myREG & mytablef.Fields("fechacrea")           'fechacrea 23
                myREG = myREG & "&" 'gion separador

            End If
           
            If mytablef.Fields("hora") = 0 Then
                myREG = myREG & Space$(10)
                myREG = myREG & "&"
            Else
                hastaCuanto = 10 - Len(mytablef.Fields("hora"))
                Call llenar_datos(hastaCuanto, Trim(mytablef.Fields("hora")), nuevoDato) 'hora 24
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If Trim(mytablef.Fields("NOMBRE")) = "vacio" Then
                myREG = myREG & Space$(80)
                myREG = myREG & "&"
            Else
                hastaCuanto = 80 - Len(mytablef.Fields("NOMBRE"))
                Call llenar_datos(hastaCuanto, Trim(mytablef.Fields("NOMBRE")), nuevoDato) ' NOMBRE 25
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("TOTAL") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("TOTAL"))
                myDato = mytablef.Fields("TOTAL")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'TOTAL 26
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If Trim$(mytablef.Fields("DESCUENTO")) = 0 Then
                myREG = myREG & Space$(7) & (0)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("DESCUENTO"))
                myDato = mytablef.Fields("DESCUENTO")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'DESCUENTO 27
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("NETO") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("NETO"))
                myDato = mytablef.Fields("NETO")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'NETO 28
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
       
            If mytablef.Fields("IMPUESTO") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("IMPUESTO"))
                myDato = mytablef.Fields("IMPUESTO")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'IMPUESTO 29
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("subtotal") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("subtotal"))
                myDato = mytablef.Fields("subtotal")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'subtotal 30
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("flage") = "?" Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                hastaCuanto = 1 - Len(mytablef.Fields("flage"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("flage"), nuevoDato) 'flage 31
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("tipo1") = "vac" Then
                myREG = myREG & Space$(3)
                myREG = myREG & "&"
            Else
                hastaCuanto = 3 - Len(mytablef.Fields("tipo1"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("tipo1"), nuevoDato) 'tipo1 32
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("serie1") = "vaci" Then
                myREG = myREG & Space$(4)
                myREG = myREG & "&"
            Else
                hastaCuanto = 4 - Len(mytablef.Fields("serie1"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("serie1"), nuevoDato) 'serie1 33
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("numero1") = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("numero1"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("numero1"), nuevoDato) 'numero1 34
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("serie2") = "vaci" Then
                myREG = myREG & Space$(4)
                myREG = myREG & "&"
            Else
                hastaCuanto = 4 - Len(mytablef.Fields("serie2"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("serie2"), nuevoDato) 'serie2 35
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("numero2") = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("numero2"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("numero2"), nuevoDato) 'numero2 36
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("serie3") = "vaci" Then
                myREG = myREG & Space$(4)
                myREG = myREG & "&"
            Else
                hastaCuanto = 4 - Len(mytablef.Fields("serie3"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("serie3"), nuevoDato) 'serie3 37
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("numero3") = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("numero3"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("numero3"), nuevoDato) 'numero3 38
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("serie4") = "vaci" Then
                myREG = myREG & Space$(4)
                myREG = myREG & "&"
            Else
                hastaCuanto = 4 - Len(mytablef.Fields("serie4"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("serie4"), nuevoDato) 'serie4 39
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("numero4") = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("numero4"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("numero4"), nuevoDato) 'numero4 40
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("serie5") = "vaci" Then
                myREG = myREG & Space$(4)
                myREG = myREG & "&"
            Else
                hastaCuanto = 4 - Len(mytablef.Fields("serie5"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("serie5"), nuevoDato) 'serie5 41
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("numero5") = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("numero5"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("numero5"), nuevoDato) 'numero5 42
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("serie6") = "vaci" Then
                myREG = myREG & Space$(4)
                myREG = myREG & "&"
            Else
                hastaCuanto = 4 - Len(mytablef.Fields("serie6"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("serie6"), nuevoDato) 'serie6 43
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("numero6") = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("numero6"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("numero6"), nuevoDato) ' numero6 44
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("serie7") = "vaci" Then
                myREG = myREG & Space$(4)
                myREG = myREG & "&"
            Else
                hastaCuanto = 4 - Len(mytablef.Fields("serie7"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("serie7"), nuevoDato) 'serie7 45
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("numero7") = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("numero7"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("numero7"), nuevoDato) 'numero7 46
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("serie8")) = "vaci" Then
                myREG = myREG & Space$(4)
                myREG = myREG & "&"
            Else
                hastaCuanto = 4 - Len(mytablef.Fields("serie8"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("serie8"), nuevoDato) 'serie8 47
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If Trim(mytablef.Fields("numero8")) = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("numero8"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("numero8"), nuevoDato) 'numero8 48
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("nop")) = "vacio" Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                hastaCuanto = 1 - Len(mytablef.Fields("nop"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("nop"), nuevoDato) 'nop 49
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
        
            If mytablef.Fields("local") = "vacio" Then
                myREG = myREG & Space$(6)
                myREG = myREG & "&"
            Else
                hastaCuanto = 6 - Trim(Len(mytablef.Fields("local")))
                myDato = Trim(mytablef.Fields("local"))
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'local 50
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If mytablef.Fields("C1") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("C1"))
                myDato = mytablef.Fields("C1")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'C1 51
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
 
            If mytablef.Fields("C2") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("C2"))
                myDato = mytablef.Fields("C2")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'C2 52
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If mytablef.Fields("C3") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("C3"))
                myDato = mytablef.Fields("C3")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'C3 53
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If mytablef.Fields("C4") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("C4"))
                myDato = mytablef.Fields("C4")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'C4 54
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If mytablef.Fields("C5") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("C5"))
                myDato = mytablef.Fields("C5")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'C5 55
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If mytablef.Fields("C6") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("C6"))
                myDato = mytablef.Fields("C6")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'C6 56
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If mytablef.Fields("C7") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("C7"))
                myDato = mytablef.Fields("C7")
                Call llenar_datos(hastaCuanto, mytablef.Fields("C7"), nuevoDato) 'C7 57
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If mytablef.Fields("C8") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("C8"))
                myDato = mytablef.Fields("C8")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'C8 59
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If mytablef.Fields("C9") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("C9"))
                myDato = mytablef.Fields("C9")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'C9 60
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
    
            If mytablef.Fields("zona") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 6 - Len(mytablef.Fields("zona"))
                myDato = mytablef.Fields("zona")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'zona 61
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("retipo1") = "vacio" Then
                myREG = myREG & Space$(6)
                myREG = myREG & "&"
            Else
                hastaCuanto = 6 - Len(mytablef.Fields("retipo1"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("retipo1"), nuevoDato) 'retipo1 62
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If mytablef.Fields("renumero1") = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("renumero1"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("renumero1"), nuevoDato) 'renumero1 63
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If mytablef.Fields("renumero2") = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("renumero2"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("renumero2"), nuevoDato) 'renumero2 64
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If mytablef.Fields("renumero3") = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("renumero3"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("renumero3"), nuevoDato) 'renumero3 65
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If mytablef.Fields("retotal") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("retotal"))
                myDato = mytablef.Fields("retotal")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'retotal 66
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
    
            If mytablef.Fields("retotal1") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("retotal1"))
                myDato = mytablef.Fields("retotal1")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'retotal1 67
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If mytablef.Fields("retotal2") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("retotal2"))
                myDato = mytablef.Fields("retotal2")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'retotal2 68
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If mytablef.Fields("retotal3") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("retotal3"))
                myDato = mytablef.Fields("retotal3")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'retotal3 69
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("acuenta") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("acuenta"))
                myDato = mytablef.Fields("acuenta")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'acuenta 70
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If mytablef.Fields("nro_items") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("nro_items"))
                myDato = mytablef.Fields("nro_items")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'nro_items 71
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If mytablef.Fields("adetotal") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("adetotal"))
                myDato = mytablef.Fields("adetotal")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'adetotal 72
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If mytablef.Fields("yausado") = "v" Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                hastaCuanto = 1 - Len(mytablef.Fields("yausado"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("yausado"), nuevoDato) 'yausado 73
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If mytablef.Fields("caja") = "vac" Then
                myREG = myREG & Space$(3)
                myREG = myREG & "&"
            Else
                hastaCuanto = 3 - Len(mytablef.Fields("caja"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("caja"), nuevoDato) 'caja 74
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If mytablef.Fields("turno") = "?" Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                hastaCuanto = 1 - Len(mytablef.Fields("turno"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("turno"), nuevoDato) 'turno 75
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If mytablef.Fields("servicio") = "?" Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                hastaCuanto = 1 - Len(mytablef.Fields("servicio"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("servicio"), nuevoDato) 'servicio 76
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If Trim(mytablef.Fields("comanda")) = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("comanda"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("comanda"), nuevoDato) 'comanda 77
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If Trim(mytablef.Fields("mesa")) = "vac" Then
                myREG = myREG & Space$(3)
                myREG = myREG & "&"
            Else
                hastaCuanto = 3 - Len(mytablef.Fields("mesa"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("mesa"), nuevoDato) 'mesa 78
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("salon")) = "vac" Then
                myREG = myREG & Space$(3)
                myREG = myREG & "&"
            Else
                hastaCuanto = 3 - Len(mytablef.Fields("salon"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("salon"), nuevoDato) 'salon 79
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("mesero")) = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("mesero"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("mesero"), nuevoDato) 'mesero 80
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If Trim(mytablef.Fields("telefono")) = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("telefono"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("telefono"), nuevoDato) 'telefono 81
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
        
            If Trim(mytablef.Fields("ruc")) = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("ruc"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("ruc"), nuevoDato) 'ruc 82
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("montopagar") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("montopagar"))
                myDato = mytablef.Fields("montopagar")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)  'montopagar 83
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If mytablef.Fields("tdocdeli") = "vacio" Then
                myREG = myREG & Space$(6)
                myREG = myREG & "&"
            Else
                hastaCuanto = 6 - Len(mytablef.Fields("tdocdeli"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("tdocdeli"), nuevoDato) 'tdocdeli 84
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("gravado")) = "" Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("gravado"))
                myDato = mytablef.Fields("gravado")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'gravado 85
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("fechasunat") = 0 Then
                myREG = myREG & Space$(10)
                myREG = myREG & "&"
            Else
                hastaCuanto = 10 - Len(mytablef.Fields("fechasunat"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("fechasunat"), nuevoDato) 'fechasunat 86
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
                
            If mytablef.Fields("flag_deli") = "?" Then
                myREG = myREG & 0
                myREG = myREG & "&"
            Else
                hastaCuanto = 1 - Len(mytablef.Fields("flag_deli"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("flag_deli"), nuevoDato) 'flag_deli 87
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("redondeo") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("redondeo"))
                myDato = mytablef.Fields("redondeo")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'redondeo 88
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("percepcion") = "" Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("percepcion"))
                myDato = mytablef.Fields("percepcion")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'percepcion 89
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If mytablef.Fields("tflete") = 0 Then
                myREG = myREG & Space$(7) & 0
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("tflete"))
                myDato = mytablef.Fields("tflete")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'tflete 90
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("localf") = "vac" Then
                myREG = myREG & Space$(2) & 0
                myREG = myREG & "&"
            Else
                hastaCuanto = 3 - Len(mytablef.Fields("localf"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("localf"), nuevoDato) 'localf 91
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("tivap") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("tivap"))
                myDato = mytablef.Fields("tivap")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'tivap 92
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("tisc") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("tisc"))
                myDato = mytablef.Fields("tisc")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'tisc 93
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
   
            If mytablef.Fields("placa") = "vacio" Then
                myREG = myREG & Space$(15)
                myREG = myREG & "&"
            Else
                hastaCuanto = 15 - Len(mytablef.Fields("placa"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("placa"), nuevoDato) 'placa 94
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("xneto") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("xneto"))
                myDato = mytablef.Fields("xneto")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'xneto 95
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("tdetra") = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("tdetra"))
                myDato = mytablef.Fields("tdetra")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'tdetra 96
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("sentido") = "?" Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                hastaCuanto = 1 - Len(mytablef.Fields("sentido"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("sentido"), nuevoDato) 'sentido 97
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("denumero") = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("denumero"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("denumero"), nuevoDato) 'denumero 98
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("dflag") = "vacio" Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                hastaCuanto = 1 - Len(mytablef.Fields("dflag"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("dflag"), nuevoDato) 'dflag 99
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
 
            If mytablef.Fields("aduana") = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("aduana"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("aduana"), nuevoDato) 'aduana 100
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("dua") = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("dua"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("dua"), nuevoDato) 'dua 101
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("importacio") = "v" Then
                myREG = myREG & Space$(1)
                myREG = myREG & "&"
            Else
                hastaCuanto = 1 - Len(mytablef.Fields("importacio"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("importacio"), nuevoDato) 'importacio 102
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If mytablef.Fields("tipoimp") = "va" Then
                myREG = myREG & Space$(2)
                myREG = myREG & "&"
            Else
                hastaCuanto = 2 - Len(mytablef.Fields("tipoimp"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("tipoimp"), nuevoDato) 'tipoimp 103
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("serieimp") = "vaci" Then
                myREG = myREG & Space$(4)
                myREG = myREG & "&"
            Else
                hastaCuanto = 4 - Len(mytablef.Fields("serieimp"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("serieimp"), nuevoDato) 'serieimp 104
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If mytablef.Fields("numeroimp") = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("numeroimp"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("numeroimp"), nuevoDato) 'numeroimp 105
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If mytablef.Fields("gasto") = "vacio" Then
                myREG = myREG & Space$(6)
                myREG = myREG & "&"
            Else
                hastaCuanto = 6 - Len(mytablef.Fields("gasto"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("gasto"), nuevoDato) 'gasto 106
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("servicioco")) = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("servicioco"))
                myDato = mytablef.Fields("servicioco")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'servicioco 107
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
     
            If Trim(mytablef.Fields("clasesunat")) = "vacio" Then
                myREG = myREG & Space$(6)
                myREG = myREG & "&"
            Else
                hastaCuanto = 6 - Len(mytablef.Fields("clasesunat"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("clasesunat"), nuevoDato) 'clasesunat 108
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If Trim(mytablef.Fields("destopo")) = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("destopo"))
                myDato = mytablef.Fields("destopo")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'destopo 109
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("horae")) = "vacio" Then
                myREG = myREG & Space$(10)
                myREG = myREG & "&"
            Else
                hastaCuanto = 10 - Len(mytablef.Fields("horae"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("horae"), nuevoDato) 'horae 110
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("vendedor1")) = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("vendedor1"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("vendedor1"), nuevoDato) 'vendedor1 111
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("vendedor2")) = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("vendedor2"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("vendedor2"), nuevoDato) 'vendedor2 112
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("vendedor3")) = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("vendedor3"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("vendedor3"), nuevoDato) 'vendedor3 113
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If Trim(mytablef.Fields("vendedor4")) = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("vendedor4"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("vendedor4"), nuevoDato) 'vendedor4 114
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If Trim(mytablef.Fields("codigo1")) = "vacio" Then
                myREG = myREG & Space$(11)
                myREG = myREG & "&"
            Else
                hastaCuanto = 11 - Len(mytablef.Fields("codigo1"))
                Call llenar_datos(hastaCuanto, mytablef.Fields("codigo1"), nuevoDato) 'codigo1 115
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("personas")) = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                hastaCuanto = 8 - Len(mytablef.Fields("personas"))
                myDato = mytablef.Fields("personas")
                Call llenar_datos(hastaCuanto, myDato, nuevoDato) 'personas 116
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            Print #Filelibero1, myREG
            Close #Filelibero1
            mytablef.MoveNext
   
            'aqui hace el progress bar
            Frm_backup.ProgressBar1.Value = ((conta / conta_record) * 100)
            Frm_backup.lblElaborandoBackup = "Elaborando backup al.." & ((conta / conta_record) * 100) & "%"
        Loop
        Close #fnum

    End If

    mytablef.Close

End Sub

Public Function read_save_factura(input_file As String, cantidad As Long)

    Dim input_record As String

    Dim my_conta     As Integer

    Dim mytablef     As New ADODB.Recordset

    mysql = ""
    fnum = FreeFile
    Open input_file For Input As #fnum

    Do Until EOF(fnum)
    
        Line Input #fnum, input_record
 
        my_conta = my_conta + 1
        'aqui llamamos a la base de datos a insertar
        mysql = "INSERT INTO factura" & Chr$(10)
        mysql = mysql & "(tipo,SERIE,NUMERO,tipoclie," & Chr$(10)
        mysql = mysql & "codigo," & Chr$(10)
        mysql = mysql & "partida,DESTINO," & Chr$(10)
        mysql = mysql & "FECHA,fechae,MONEDA,VENDEDOR,transporte," & Chr$(10)
        mysql = mysql & "fpago," & Chr$(10)
        mysql = mysql & "paridad," & Chr$(10)
        mysql = mysql & "dias," & Chr$(10)
        mysql = mysql & "bodega,bodegaf," & Chr$(10)
        mysql = mysql & "observa,estado," & Chr$(10)
        mysql = mysql & "acu,acu1,usuario,fechacrea,hora," & Chr$(10)
        mysql = mysql & "nombre,TOTAL,DESCUENTO,NETO,IMPUESTO,SUBTOTAL," & Chr$(10)
        mysql = mysql & "flage,TIPO1,SERIE1,NUMERO1," & Chr$(10)
        mysql = mysql & "SERIE2,NUMERO2,SERIE3,NUMERO3," & Chr$(10)
        mysql = mysql & "SERIE4,NUMERO4,SERIE5,NUMERO5," & Chr$(10)
        mysql = mysql & "SERIE6,NUMERO6," & Chr$(10)
        mysql = mysql & "SERIE7,NUMERO7," & Chr$(10)
        mysql = mysql & "SERIE8,NUMERO8," & Chr$(10)
        mysql = mysql & "NOP,LOCAL," & Chr$(10)
        mysql = mysql & "C1,C2,C3,C4,C5,C6,C7,C8,C9," & Chr$(10)
        mysql = mysql & "zona,retipo1,renumero1,renumero2,renumero3," & Chr$(10)
        mysql = mysql & "retotal,retotal1,retotal2,retotal3," & Chr$(10)
        mysql = mysql & "acuenta,nro_items,adetotal," & Chr$(10)
        mysql = mysql & "yausado,caja,turno,servicio," & Chr$(10)
        mysql = mysql & "comanda,mesa,salon,mesero,telefono,ruc," & Chr$(10)
        mysql = mysql & "montopagar,tdocdeli,gravado,fechasunat," & Chr$(10)
        mysql = mysql & "flag_deli,redondeo,percepcion," & Chr$(10)
        mysql = mysql & "tflete,localf,tivap," & Chr$(10)
        mysql = mysql & "tisc,placa,xneto,tdetra," & Chr$(10)
        mysql = mysql & "sentido,denumero,dflag,aduana," & Chr$(10)
        mysql = mysql & "dua,importacio,tipoimp," & Chr$(10)
        mysql = mysql & "serieimp,numeroimp,gasto,servicioco," & Chr$(10)
        mysql = mysql & "clasesunat,destopo,horae," & Chr$(10)
        mysql = mysql & "vendedor1,vendedor2,vendedor3," & Chr$(10)
        mysql = mysql & "vendedor4,codigo1,personas)" & Chr$(10)
        mysql = mysql & " VALUES ('" & Mid(input_record, 1, 3) & "'," & Chr$(10) 'tipo 1
        mysql = mysql & " '" & Mid(input_record, 5, 4) & "'," & Chr$(10) 'SERIE 2
        mysql = mysql & " '" & Trim(Mid(input_record, 10, 11)) & "'," & Chr$(10) 'NUMERO 3
        mysql = mysql & " '" & Trim(Mid(input_record, 22, 1)) & "'," & Chr$(10) 'tipoclie 4
        mysql = mysql & " '" & Trim(Mid(input_record, 24, 11)) & "'," & Chr$(10) 'codigo 5
        mysql = mysql & " '" & Trim$((Mid(input_record, 36, 60))) & "'," & Chr$(10) 'PARTIDA 6
        mysql = mysql & " '" & Trim(Mid(input_record, 97, 60)) & "'," & Chr$(10) 'DESTINO 7
        mysql = mysql & " '" & Trim$((Mid(input_record, 158, 10))) & "'," & Chr$(10) 'FECHA 8
        mysql = mysql & " '" & Trim$((Mid(input_record, 169, 10))) & "'," & Chr$(10) 'fechae 9
        mysql = mysql & " '" & Trim$((Mid(input_record, 180, 1))) & "'," & Chr$(10) 'MONEDA 10
        mysql = mysql & " '" & Trim$((Mid(input_record, 182, 11))) & "'," & Chr$(10) 'VENDEDOR 11

        If (Mid(input_record, 194, 11)) = Space(11) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & " '" & Trim$((Mid(input_record, 194, 11))) & "'," & Chr$(10) 'transporte 12

        End If

        mysql = mysql & " '" & Trim$((Mid(input_record, 206, 3))) & "'," & Chr$(10) 'FPAGO 13
        mysql = mysql & " '" & Trim$((Mid(input_record, 210, 8))) & "'," & Chr$(10) 'PARIDAD 14
        mysql = mysql & " '" & Trim((Mid(input_record, 219, 8))) & "'," & Chr$(10) 'DIAS 15
        mysql = mysql & " '" & Trim$(Mid(input_record, 228, 2)) & "'," & Chr$(10) 'BODEGA 16
        mysql = mysql & " '" & Trim$(Mid(input_record, 231, 2)) & "'," & Chr$(10) 'bodegaf 17
        mysql = mysql & " '" & Trim$(Mid(input_record, 234, 60)) & "'," & Chr$(10) 'observa 18
        mysql = mysql & " '" & Trim$(Mid(input_record, 295, 1)) & "'," & Chr$(10) 'estado 19
        mysql = mysql & " '" & Trim$(Mid(input_record, 297, 1)) & "'," & Chr$(10) 'acu 20
        mysql = mysql & " '" & Trim$(Mid(input_record, 299, 1)) & "'," & Chr$(10) 'acu1 21
        mysql = mysql & " '" & Trim$(Mid(input_record, 301, 11)) & "'," & Chr$(10) 'usuario 22
        mysql = mysql & " '" & Trim$(Mid(input_record, 313, 10)) & "'," & Chr$(10) 'fechacrea 23
        mysql = mysql & " '" & Trim$((Mid(input_record, 324, 10))) & "'," & Chr$(10) 'hora 24
        mysql = mysql & " '" & Trim$(Mid(input_record, 335, 80)) & "'," & Chr$(10) 'NOMBRE 25

        If (Mid(input_record, 416, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 416, 8)) & "," & Chr$(10) 'TOTAL 26

        End If

        If (Mid(input_record, 425, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 425, 8)) & "," & Chr$(10) 'DESCUENTO 27

        End If

        If (Mid(input_record, 434, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 434, 8)) & "," & Chr$(10) 'NETO 28

        End If

        If (Mid(input_record, 443, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 443, 8)) & "," & Chr$(10) 'IMPUESTO 29

        End If

        '    mysql = mysql & "" & (Mid(input_record, 443, 8)) & "," & Chr$(10) 'SUBTOTAL 30
        mysql = mysql & "'" & Trim$(Mid(input_record, 452, 8)) & "'," & Chr$(10) 'FLAGE 31
        mysql = mysql & "'" & Trim$(Mid(input_record, 461, 1)) & "'," & Chr$(10) 'TIPO1 32
        mysql = mysql & "'" & Trim$(Mid(input_record, 463, 3)) & "'," & Chr$(10) 'SERIE1 33
        mysql = mysql & "'" & Trim$(Mid(input_record, 467, 4)) & "'," & Chr$(10) 'NUMERO1 34
        mysql = mysql & "'" & Trim$(Mid(input_record, 472, 11)) & "'," & Chr$(10) 'SERIE2 35
        mysql = mysql & "'" & Trim$(Mid(input_record, 484, 4)) & "'," & Chr$(10) 'NUMERO2 36
        mysql = mysql & "'" & Trim$(Mid(input_record, 489, 11)) & "'," & Chr$(10) 'SERIE3 37
        mysql = mysql & "'" & Trim$(Mid(input_record, 501, 4)) & "'," & Chr$(10) 'NUMERO3 38
        mysql = mysql & "'" & Trim$(Mid(input_record, 506, 11)) & "'," & Chr$(10) 'SERIE4 39
        mysql = mysql & "'" & Trim$(Mid(input_record, 518, 4)) & "'," & Chr$(10) 'NUMERO4 40
        mysql = mysql & "'" & Trim$(Mid(input_record, 523, 11)) & "'," & Chr$(10) 'SERIE5 41
        mysql = mysql & "'" & Trim$(Mid(input_record, 535, 4)) & "'," & Chr$(10) 'NUMERO5 41
        mysql = mysql & "'" & Trim$(Mid(input_record, 540, 11)) & "'," & Chr$(10) 'SERIE6 42
        mysql = mysql & "'" & Trim$(Mid(input_record, 552, 4)) & "'," & Chr$(10) 'NUMERO6 43
        mysql = mysql & "'" & Trim$(Mid(input_record, 557, 11)) & "'," & Chr$(10) 'SERIE7 44
        mysql = mysql & "'" & Trim$(Mid(input_record, 569, 4)) & "'," & Chr$(10) 'NUMERO7 45
        mysql = mysql & "'" & Trim$(Mid(input_record, 574, 11)) & "'," & Chr$(10) 'SERIE8 46
        mysql = mysql & "'" & Trim$(Mid(input_record, 586, 4)) & "'," & Chr$(10) 'NUMERO8 47
    
        mysql = mysql & " '" & Trim$(Mid(input_record, 591, 11)) & "'," & Chr$(10)
    
        mysql = mysql & " '" & Trim$((Mid(input_record, 603, 1))) & "'," & Chr$(10) 'NOP 48
    
        If (Mid(input_record, 605, 6)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 605, 6)) & "'," & Chr$(10) 'local 49

        End If
    
        If (Mid(input_record, 612, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 612, 8)) & "," & Chr$(10) 'C2 51

        End If

        If (Mid(input_record, 621, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 621, 8)) & "," & Chr$(10) 'C3 52

        End If

        If (Mid(input_record, 630, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 630, 8)) & "," & Chr$(10) 'C4 53

        End If

        If (Mid(input_record, 639, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 639, 8)) & "," & Chr$(10) 'C5 54

        End If

        If (Mid(input_record, 648, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 648, 8)) & "," & Chr$(10) 'C6 55

        End If

        If (Mid(input_record, 657, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 657, 8)) & "," & Chr$(10) 'C7 56

        End If
    
        If (Mid(input_record, 666, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 666, 8)) & "," & Chr$(10) 'C8 57

        End If

        If (Mid(input_record, 675, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 675, 8)) & "," & Chr$(10) 'C9 58

        End If

        If (Mid(input_record, 684, 6)) = Space$(6) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 684, 6)) & "'," & Chr$(10) 'zona 59

        End If
    
        mysql = mysql & "'" & Trim$(Mid(input_record, 693, 6)) & "'," & Chr$(10) 'retipo1 60
        mysql = mysql & "'" & Trim$(Mid(input_record, 700, 6)) & "'," & Chr$(10) 'renumero1 61
        mysql = mysql & "'" & Trim$(Mid(input_record, 707, 11)) & "'," & Chr$(10) 'renumero2 62
        mysql = mysql & "'" & Trim$(Mid(input_record, 719, 11)) & "'," & Chr$(10) 'renumero3 63

        If (Mid(input_record, 731, 11)) = Space(11) Then
            mysql = mysql & 0 & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 731, 11)) & "," & Chr$(10) 'retotal 64

        End If
    
        If (Mid(input_record, 743, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 743, 8)) & "," & Chr$(10) 'retotal1 65

        End If

        If (Mid(input_record, 752, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 752, 8)) & "," & Chr$(10) 'retotal2 66

        End If

        If (Mid(input_record, 761, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 761, 8)) & "," & Chr$(10) 'retotal3 67

        End If

        If (Mid(input_record, 770, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 770, 8)) & "," & Chr$(10) 'acuenta 68

        End If

        If (Mid(input_record, 779, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 779, 8)) & "," & Chr$(10) 'nro_items 69

        End If

        If (Mid(input_record, 788, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 788, 8)) & "," & Chr$(10) 'adetotal 70

        End If

        mysql = mysql & "'" & Trim$(Mid(input_record, 797, 8)) & "'," & Chr$(10) 'yausado 71
        mysql = mysql & "'" & Trim$(Mid(input_record, 806, 1)) & "'," & Chr$(10) 'caja 72
        mysql = mysql & "'" & Trim$(Mid(input_record, 808, 3)) & "'," & Chr$(10) 'turno 73
        mysql = mysql & "'" & Trim$(Mid(input_record, 812, 1)) & "'," & Chr$(10) 'servicio 74
        mysql = mysql & "'" & Trim$(Mid(input_record, 814, 1)) & "'," & Chr$(10) 'comanda 75
        mysql = mysql & "'" & Trim$(Mid(input_record, 816, 11)) & "'," & Chr$(10) 'mesa 76
        mysql = mysql & "'" & Trim$(Mid(input_record, 828, 3)) & "'," & Chr$(10) 'salon 77
        mysql = mysql & "'" & Trim$(Mid(input_record, 832, 3)) & "'," & Chr$(10) 'mesero 78
        mysql = mysql & "'" & Trim$(Mid(input_record, 836, 11)) & "'," & Chr$(10) 'telefono 79
        mysql = mysql & "'" & Trim$(Mid(input_record, 848, 11)) & "'," & Chr$(10) 'ruc 80
        mysql = mysql & "'" & Trim$(Mid(input_record, 860, 11)) & "'," & Chr$(10) 'montopagar 81
        mysql = mysql & "'" & Trim$(Mid(input_record, 872, 8)) & "'," & Chr$(10) 'tdocdeli 82

        If (Mid(input_record, 881, 6)) = Space$(6) Then
            mysql = mysql & 0 & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 881, 6)) & "," & Chr$(10) 'gravado 83

        End If
    
        If (Mid(input_record, 888, 8)) = Space(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 888, 8)) & "," & Chr$(10) 'fechasunat 84

        End If

        If (Mid(input_record, 897, 10)) = Space$(10) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 897, 10)) & "," & Chr$(10) 'flag_deli '85

        End If

        mysql = mysql & "" & (Mid(input_record, 908, 1)) & "," & Chr$(10) ' 86

        If (Mid(input_record, 910, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 910, 8)) & "," & Chr$(10) 'redondeo 87

        End If
    
        If (Mid(input_record, 919, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 919, 8)) & "," & Chr$(10) 'gravado 88

        End If

        If (Mid(input_record, 928, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 928, 8)) & "," & Chr$(10) 'tflete 89

        End If

        If (Mid(input_record, 937, 3)) = Space$(3) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 937, 3)) & "," & Chr$(10) 'localf 90

        End If

        If (Mid(input_record, 941, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 941, 8)) & "," & Chr$(10) 'tivap 91

        End If
    
        If (Mid(input_record, 950, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 950, 8)) & "," & Chr$(10) 'tisc 92

        End If

        '
        If (Mid(input_record, 959, 15)) = Space(15) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 959, 15)) & "'," & Chr$(10) 'placa 93

        End If

        If (Mid(input_record, 975, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 975, 8)) & "," & Chr$(10) 'xneto 94

        End If

        If (Mid(input_record, 984, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "" & Trim$(Mid(input_record, 984, 8)) & "," & Chr$(10) 'tdetra 95

        End If
    
        If (Mid(input_record, 993, 1)) = Space$(1) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 993, 1)) & "'," & Chr$(10) 'sentido 96

        End If
    
        mysql = mysql & "'" & Trim$(Mid(input_record, 995, 11)) & "'," & Chr$(10) 'denumero 97

        If (Mid(input_record, 1007, 1)) = Space(1) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1007, 1)) & "'," & Chr$(10) 'dflag 98

        End If

        If (Mid(input_record, 1009, 11)) = Space(11) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1009, 11)) & "'," & Chr$(10) 'aduana 99

        End If

        '
        If (Mid(input_record, 1021, 11)) = Space(11) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1021, 11)) & "'," & Chr$(10) 'dua 100

        End If

        If (Mid(input_record, 1033, 1)) = Space(1) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1033, 1)) & "'," & Chr$(10) 'importacio 101

        End If

        If (Mid(input_record, 1035, 2)) = Space(2) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1035, 2)) & "'," & Chr$(10) 'tipoimp 102

        End If

        If (Mid(input_record, 1038, 4)) = Space(4) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1038, 4)) & "'," & Chr$(10) 'serieimp 103

        End If

        If (Mid(input_record, 1043, 11)) = Space(11) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1043, 11)) & "'," & Chr$(10) 'numeroimp 104

        End If

        If (Mid(input_record, 1055, 6)) = Space(6) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1055, 6)) & "'," & Chr$(10) 'gasto 105

        End If

        If (Mid(input_record, 1062, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1062, 8)) & "'," & Chr$(10) 'servicioco 106

        End If

        If (Mid(input_record, 1071, 8)) = Space$(6) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1071, 8)) & "'," & Chr$(10) 'clasesunat 107

        End If

        If (Mid(input_record, 1078, 6)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1078, 6)) & "'," & Chr$(10) 'destopo 108

        End If

        mysql = mysql & "'" & Trim$(Mid(input_record, 1087, 10)) & "'," & Chr$(10) 'horae 109

        If (Mid(input_record, 1098, 11)) = Space$(11) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1098, 11)) & "'," & Chr$(10) 'vendedor2 110

        End If

        If (Mid(input_record, 1110, 11)) = Space$(11) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1110, 11)) & "'," & Chr$(10) 'vendedor3 111

        End If

        If (Mid(input_record, 1122, 11)) = Space$(11) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1122, 11)) & "'," & Chr$(10) 'vendedor4 112

        End If

        If (Mid(input_record, 1134, 11)) = Space$(11) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1134, 11)) & "'," & Chr$(10) 'codigo1 113

        End If

        If (Mid(input_record, 1146, 8)) = Space$(8) Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1146, 8)) & "'," & Chr$(10) 'personas 114

        End If

        If (Mid(input_record, 1158, 8)) = Space$(8) Then
            mysql = mysql & "Null" & ")" & Chr$(10)
        Else
            mysql = mysql & "'" & Trim$(Mid(input_record, 1158, 8)) & "')" & Chr$(10) 'personas 115

        End If

        cn.Execute (mysql)
  
        If cantidad < my_conta Then
            Exit Do

        End If

        Frm_backup.ProgressBar1.Value = ((my_conta / cantidad) * 100)
        Frm_backup.lblElaborandoBackup.Caption = "Actualizando al.." & ((my_conta / cantidad) * 100) & "%"
        '    mytablef.MoveLast

    Loop
    Close #fnum

End Function

'fino 02/05/2017 pll

'inicio 04/05/2017 pll

Public Sub bkp_almacen()

    Dim mysql       As String

    Dim mytablef    As New ADODB.Recordset

    Dim hastaCuanto As Integer

    Dim myDato      As String

    Dim nuevoDato   As String

    Dim my_Lnumero  As Integer

    Dim conta       As Integer

    FileName = "C:\EmpaquetaVi\" & "bkpAlmacen" & ".txt"

    mysql = "SELECT isnull(producto,'vacio') as producto," & Chr$(10)
    mysql = mysql & "isnull(local,'vacio') as local, " & Chr$(10)
    mysql = mysql & "isnull(bodega,'vacio') as bodega, " & Chr$(10)
    mysql = mysql & "isnull(saldo,0) as saldo, " & Chr$(10)
    mysql = mysql & "isnull(t1,0) as t1, " & Chr$(10)
    mysql = mysql & "isnull(t2,0) as t2, " & Chr$(10)
    mysql = mysql & "isnull(t3,0) as t3," & Chr$(10)
    mysql = mysql & "isnull(t4,0) as t4," & Chr$(10)
    mysql = mysql & "isnull(t5,0) as t5," & Chr$(10)
    mysql = mysql & "isnull(t6,0) as t6," & Chr$(10)
    mysql = mysql & "isnull(t7,0) as t7," & Chr$(10)
    mysql = mysql & "isnull(t8,0) as t8," & Chr$(10)
    mysql = mysql & "isnull(t9,0) as t9," & Chr$(10)
    mysql = mysql & "isnull(t10,0) as t10," & Chr$(10)
    mysql = mysql & "isnull(t11,0) as t11," & Chr$(10)
    mysql = mysql & "isnull(t12,0) as t12," & Chr$(10)
    mysql = mysql & "isnull(t13,0) as t13," & Chr$(10)
    mysql = mysql & "isnull(t14,0) as t14," & Chr$(10)
    mysql = mysql & "isnull(t15,0) as t15," & Chr$(10)
    mysql = mysql & "isnull(t16,0) as t16," & Chr$(10)
    mysql = mysql & "isnull(minimo,0) as minimo," & Chr$(10)
    mysql = mysql & "isnull(maximo,0) as maximo," & Chr$(10)
    mysql = mysql & "isnull(ccosto,0) as ccosto," & Chr$(10)
    mysql = mysql & "isnull(entrada,0) as entrada," & Chr$(10)
    mysql = mysql & "isnull(salida,0) salida," & Chr$(10)
    mysql = mysql & "isnull(saldoinicial,0) as saldoinicial," & Chr$(10)
    mysql = mysql & "isnull(unidad,'vacio') as unidad," & Chr$(10)
    mysql = mysql & "isnull(costo,0) as costo" & Chr$(10)
    mysql = mysql & "from almacen " & Chr$(10)
    'mysql = mysql & "order by producto" & Chr$(10)

    mytablef.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablef.RecordCount > 0 Then
        Do

            If mytablef.EOF Then Exit Do
            'aqui es para abrir el file txt y llenarlo
            Filelibero = FreeFile
            Open FileName For Append As #Filelibero
            conta = conta + 1
      
            conta_record = mytablef.RecordCount
            myREG = ""

            If Trim(mytablef.Fields("producto")) = "vacio" Then
                myREG = myREG & Space$(15)
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("producto"))
                myDato = mytablef.Fields("producto")
                hastaCuanto = 15 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If Trim(mytablef.Fields("local")) = "vacio" Then
                myREG = myREG & Space$(10)
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("local"))
                myDato = mytablef.Fields("local")
                hastaCuanto = 10 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("bodega")) = "va" Then
                myREG = myREG & Space$(2)
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("bodega"))
                myDato = mytablef.Fields("bodega")
                hastaCuanto = 2 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
  
            If Trim(mytablef.Fields("saldo")) = 0 Then
                myREG = myREG & Space$(10)
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("saldo"))
                myDato = mytablef.Fields("saldo")
                hastaCuanto = 10 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("t1")) = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("t1"))
                myDato = mytablef.Fields("t1")
                hastaCuanto = 8 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("t2")) = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("t2"))
                myDato = mytablef.Fields("t2")
                hastaCuanto = 8 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("t3")) = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("t3"))
                myDato = mytablef.Fields("t3")
                hastaCuanto = 8 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("t4")) = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("t4"))
                myDato = mytablef.Fields("t4")
                hastaCuanto = 8 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
 
            If Trim(mytablef.Fields("t5")) = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("t5"))
                myDato = mytablef.Fields("t5")
                hastaCuanto = 8 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("t6")) = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("t6"))
                myDato = mytablef.Fields("t6")
                hastaCuanto = 8 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("t7")) = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("t7"))
                myDato = mytablef.Fields("t7")
                hastaCuanto = 8 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("t8")) = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("t8"))
                myDato = mytablef.Fields("t8")
                hastaCuanto = 8 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("t9")) = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("t9"))
                myDato = mytablef.Fields("t9")
                hastaCuanto = 8 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("t10")) = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("t10"))
                myDato = mytablef.Fields("t10")
                hastaCuanto = 8 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("t11")) = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("t11"))
                myDato = mytablef.Fields("t11")
                hastaCuanto = 8 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("t12")) = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("t12"))
                myDato = mytablef.Fields("t12")
                hastaCuanto = 8 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If
      
            If Trim(mytablef.Fields("t13")) = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("t13"))
                myDato = mytablef.Fields("t13")
                hastaCuanto = 8 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("t14")) = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("t14"))
                myDato = mytablef.Fields("t14")
                hastaCuanto = 8 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("t15")) = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("t15"))
                myDato = mytablef.Fields("t15")
                hastaCuanto = 8 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("t16")) = 0 Then
                myREG = myREG & Space$(8)
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("t16"))
                myDato = mytablef.Fields("t16")
                hastaCuanto = 8 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("minimo")) = 0 Then
                myREG = myREG & Space$(7) & 0
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("minimo"))
                myDato = mytablef.Fields("minimo")
                hastaCuanto = 8 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("maximo")) = 0 Then
                myREG = myREG & Space$(7) & 0
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("maximo"))
                myDato = mytablef.Fields("maximo")
                hastaCuanto = 8 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("ccosto")) = 0 Then
                myREG = myREG & Space$(5) & 0
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("ccosto"))
                myDato = mytablef.Fields("ccosto")
                hastaCuanto = 6 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("entrada")) = 0 Then
                myREG = myREG & Space$(7) & 0
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("entrada"))
                myDato = mytablef.Fields("entrada")
                hastaCuanto = 8 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("salida")) = 0 Then
                myREG = myREG & Space$(7) & 0
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("salida"))
                myDato = mytablef.Fields("salida")
                hastaCuanto = 8 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("saldoinicial")) = 0 Then
                myREG = myREG & Space$(7) & 0
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("saldoinicial"))
                myDato = mytablef.Fields("saldoinicial")
                hastaCuanto = 8 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("unidad")) = "vacio" Then
                myREG = myREG & Space$(5) & 0
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("unidad"))
                myDato = mytablef.Fields("unidad")
                hastaCuanto = 6 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            If Trim(mytablef.Fields("costo")) = 0 Then
                myREG = myREG & Space$(7) & 0
                myREG = myREG & "&"
            Else
                my_Lnumero = Len(mytablef.Fields("costo"))
                myDato = mytablef.Fields("costo")
                hastaCuanto = 8 - my_Lnumero
                Call llenar_datos(hastaCuanto, myDato, nuevoDato)
                myREG = myREG & nuevoDato
                myREG = myREG & "&"

            End If

            Print #Filelibero, myREG
            Close #Filelibero
            mytablef.MoveNext
   
            'aqui hace el progress bar
            Frm_backup.ProgressBar1.Value = ((conta / conta_record) * 100)
            Frm_backup.lblElaborandoBackup = "Elaborando backup al.." & ((conta / conta_record) * 100) & "%"
        Loop
        Close #Filelibero

    End If

    mytablef.Close

End Sub

Public Function read_save_almacen(input_file As String, cantidad As Long)

    Dim input_record As String

    Dim my_conta     As Integer

    Dim mytablef     As New ADODB.Recordset

    fnum = FreeFile
    Open input_file For Input As #fnum

    Do Until EOF(fnum)
    
        Line Input #fnum, input_record
 
        my_conta = my_conta + 1
        'aqui llamamos a la base de datos a insertar
        mysql = "INSERT INTO almacen " & Chr$(10)
        mysql = mysql & "(producto,local,bodega," & Chr$(10)
        mysql = mysql & "saldo,t1,t2," & Chr$(10)
        mysql = mysql & "t3,t4," & Chr$(10)
        mysql = mysql & "t5,t6,t7,t8," & Chr$(10)
        mysql = mysql & "t9,t10,t11,t12,t13," & Chr$(10)
        mysql = mysql & "t14,t15,t16," & Chr$(10)
        mysql = mysql & "minimo,maximo," & Chr$(10)
        mysql = mysql & "ccosto,entrada,salida," & Chr$(10)
        mysql = mysql & "saldoinicial,unidad,costo)" & Chr$(10)

        mysql = mysql & " VALUES ('" & Trim(Mid(input_record, 1, 15)) & "'," & Chr$(10) 'producto
        mysql = mysql & " '" & Trim(Mid(input_record, 17, 10)) & "'," & Chr$(10) 'local
        mysql = mysql & " '" & Trim(Mid(input_record, 28, 2)) & "'," & Chr$(10) 'bodega
        mysql = mysql & " '" & Trim(Mid(input_record, 31, 10)) & "'," & Chr$(10) 'saldo
        mysql = mysql & " '" & Trim(Mid(input_record, 42, 8)) & "'," & Chr$(10) 't1
        mysql = mysql & " '" & Trim(Mid(input_record, 51, 8)) & "'," & Chr$(10) 't2
        mysql = mysql & " '" & Trim(Mid(input_record, 60, 8)) & "'," & Chr$(10) 't3
        mysql = mysql & " '" & Trim(Mid(input_record, 69, 8)) & "'," & Chr$(10) 't4
        mysql = mysql & " '" & Trim(Mid(input_record, 78, 8)) & "'," & Chr$(10) 't5
        mysql = mysql & " '" & Trim(Mid(input_record, 87, 8)) & "'," & Chr$(10) 't6
        mysql = mysql & " '" & Trim(Mid(input_record, 96, 8)) & "'," & Chr$(10) 't7
        mysql = mysql & " '" & Trim(Mid(input_record, 105, 8)) & "'," & Chr$(10) 't8
        mysql = mysql & " '" & Trim(Mid(input_record, 114, 8)) & "'," & Chr$(10) 't9
        mysql = mysql & " '" & Trim(Mid(input_record, 123, 8)) & "'," & Chr$(10) 't10
        mysql = mysql & " '" & Trim(Mid(input_record, 132, 8)) & "'," & Chr$(10) 't11
        mysql = mysql & " '" & Trim(Mid(input_record, 141, 8)) & "'," & Chr$(10) 't12
        mysql = mysql & " '" & Trim(Mid(input_record, 150, 8)) & "'," & Chr$(10) 't13
        mysql = mysql & " '" & Trim(Mid(input_record, 159, 8)) & "'," & Chr$(10) 't14
        mysql = mysql & " '" & Trim(Mid(input_record, 168, 8)) & "'," & Chr$(10) 't15
        mysql = mysql & " '" & Trim(Mid(input_record, 177, 8)) & "'," & Chr$(10) 't16
    
        If Trim(Mid(input_record, 186, 8)) = 0 Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & " '" & Trim(Mid(input_record, 186, 8)) & "'," & Chr$(10) 'minimo

        End If
    
        If Trim(Mid(input_record, 195, 8)) = 0 Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & " '" & Trim(Mid(input_record, 195, 8)) & "'," & Chr$(10) 'maximo

        End If

        If Trim(Mid(input_record, 204, 6)) = 0 Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & " '" & Trim(Mid(input_record, 204, 6)) & "'," & Chr$(10) 'ccostos

        End If
    
        If Trim(Mid(input_record, 211, 8)) = 0 Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & " '" & Trim(Mid(input_record, 211, 8)) & "'," & Chr$(10) 'entrada

        End If
    
        If Trim(Mid(input_record, 220, 8)) = 0 Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & " '" & Trim(Mid(input_record, 220, 8)) & "'," & Chr$(10) 'salida

        End If

        If Trim(Mid(input_record, 229, 8)) = 0 Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & " '" & Trim(Mid(input_record, 229, 8)) & "'," & Chr$(10) 'saldoinicial

        End If
    
        If Trim(Mid(input_record, 238, 6)) = 0 Then
            mysql = mysql & "Null" & "," & Chr$(10)
        Else
            mysql = mysql & " '" & Trim(Mid(input_record, 238, 6)) & "'," & Chr$(10) 'unidad

        End If
    
        If Trim(Mid(input_record, 245, 8)) = 0 Then
            mysql = mysql & "Null" & ")" & Chr$(10)
        Else
            mysql = mysql & " '" & Trim(Mid(input_record, 245, 8)) & "')" & Chr$(10) 'costo

        End If
    
        cn.Execute (mysql)
  
        If cantidad < my_conta Then
            Exit Do

        End If

        Frm_backup.ProgressBar1.Value = ((my_conta / cantidad) * 100)
        Frm_backup.lblElaborandoBackup.Caption = "Actualizando al.." & ((my_conta / cantidad) * 100) & "%"
        '    mytablef.MoveLast

    Loop
    Close #fnum

End Function

'fin 04/05/2017 pll
Public Sub crear_rar(finicio As String, ffinal As String)

    Dim carpetaToExtract As String

    carpetaToExtract = "C:\EmpaquetaVi"
    creafilerar = "C:\EmpaquetaVi" & finicio & ffinal & ".rar"
    Shell "C:\Program Files\WinRAR\WinRAR.exe a " & creafilerar & " " & carpetaToExtract, vbHide

End Sub

Public Sub desampaqueta_rar()

    Dim fileToExtract     As String

    Dim destinationBackup As String

    Dim crearCarpeta      As String

    Dim mydestino         As String

    On Error GoTo desampaqueta_err

    origenDestination = "C:\" & myfile
    MkDir ("C:\DesempaquetaVi") '
    mydestino = "C:\DesempaquetaVi\"
 
    Shell "C:\Program Files\WinRAR\WinRAR.exe e " & origenDestination & " " & mydestino, vbHide
   
desampaqueta_err:
    Exit Sub

End Sub

Public Sub copia_Decargar()

    Dim destino As String

    destino = "C:\" & myfile

    FileCopy origen, destino

End Sub

Public Sub crea_directorio()

    On Error GoTo crea_error

    MkDir ("C:\EmpaquetaVi")
crea_error:
    Exit Sub

End Sub

Public Function control_factura(finicio, ffinal, mmyprocedef)

    Dim mysql    As String

    Dim mytablef As New ADODB.Recordset

    mysql = "SELECT * from factura where fecha>='" & finicio & "'" & Chr$(10)
    mysql = mysql & "and fecha<='" & ffinal & "' " & Chr$(10)
  
    mytablef.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablef.RecordCount = 0 Then
        myprocedef = True
    Else
        myprocedef = False
       
    End If

End Function

Public Function Eli_Bckp_detalle(fechai As String, fechaf As String)

    Dim mysql      As String

    Dim mytablef   As New ADODB.Recordset

    Dim myproceded As Boolean

    mysql = "Delete " & Chr$(10)
    mysql = mysql & "from detalle where fecha>='" & fechai & "'" & Chr$(10)
    mysql = mysql & "and fecha<='" & fechaf & "' " & Chr$(10)

    mytablef.Open mysql, cn, adOpenStatic, adLockOptimistic
   
    cn.Execute (mysql)
    'cn.Close
   
End Function

Public Function Eli_Bckp_factura(fechai As String, ffinal As String)

    Dim mysql      As String

    Dim mytablef   As New ADODB.Recordset

    Dim myproceded As Boolean

    mysql = "Delete " & Chr$(10)
    mysql = mysql & "from factura where fecha>='" & fechai & "'" & Chr$(10)
    mysql = mysql & "and fecha<='" & ffinal & "' " & Chr$(10)

    mytablef.Open mysql, cn, adOpenStatic, adLockOptimistic
   
    cn.Execute (mysql)

    ' cn.Close
End Function

Public Sub control_detalle(finicio, ffinal, myproceded)

    Dim mysql    As String

    Dim mytablef As New ADODB.Recordset

    mysql = "SELECT * from detalle where fecha>='" & finicio & "'" & Chr$(10)
    mysql = mysql & "and fecha<='" & ffinal & "' " & Chr$(10)

    mytablef.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytablef.RecordCount = 0 Then
        myproceded = True
    Else
        myproceded = False

    End If

End Sub

Public Function Eli_Bckp_almacen()

    Dim mysql      As String

    Dim mytablef   As New ADODB.Recordset

    Dim myproceded As Boolean

    mysql = "Delete " & Chr$(10)
    mysql = mysql & "from almacen" & Chr$(10)

    mytablef.Open mysql, cn, adOpenStatic, adLockOptimistic
   
    cn.Execute (mysql)

    ' cn.Close
End Function

Public Sub enviar_correoBck()

    Dim mytablex As New ADODB.Recordset

    mysql = "select * " & Chr$(10)
    mysql = mysql & "from correos " & Chr$(10)
     
    mytablef.Open mysql, cn, adOpenStatic, adLockOptimistic
     
    myserver = Trim("" & mytablex.Fields("txtserver"))
    myusername = Trim("" & mytablex.Fields("txtusername"))
    mypassword = Trim("" & mytablex.Fields("txtpassword"))
    myfromname = Trim("" & mytablex.Fields("txtfromname"))
    myfromemail = Trim("" & mytablex.Fields("txtfromemail"))
    myport = Trim("" & mytablex.Fields("txtport"))
    myselecciona = Trim("" & mytablex.Fields("txtselecciona"))
    mytto = Trim("" & mytablex.Fields("txtto"))
    chkssl = Trim("" & mytablex.Fields("chkssl"))
    myattach = "C:\" & creafilerar
    mysubject = Trim("" & mytablex.Fields("txtsubject"))
    mymsg = Trim("" & mytablex.Fields("txtmsg"))
    mymsg = txtmsg & Chr$(10) & Chr$(13) & ""
    mymsg = txtmsg & Format(Now, "dd/mm/yyyy") + " " + Format(Now, "hh:mm:ss")
 
    retval = SendMail(Trim$(myfromemail), Trim$(mysubject), Trim(myfromname) & "<" & Trim(myfromemail) & ">", Trim$(mymsg), Trim(myserver), CInt(txtport), Trim(myusername), Trim(mypassword), Trim(myattach), True, Trim(myselecciona), Trim(myhtml))
          
    If retval = 1 Then
        MsgBox "Proceso Realizado ", 48, "Aviso"

    End If
 
    mytablex.Close

End Sub

Sub envio_correosBackup()

    Dim txtserver     As String

    Dim txtusername   As String

    Dim txtpassword   As String

    Dim txtport       As String

    Dim txtto         As String

    Dim chkssl        As String

    Dim txtfromname   As String

    Dim txtfromemail  As String

    Dim txtattach     As String

    Dim txtsubject    As String

    Dim txtmsg        As String

    Dim retval        As String

    Dim txthtml       As String

    Dim txtselecciona As String

    'Dim txtselecciona As String
    Dim mytablex      As New ADODB.Recordset

    Dim buf           As String

    On Error GoTo cmd0905677_err

    mytablex.Open "select * from correos where cosms='11'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        MsgBox "Correo No Configurado", vbCritical, "Message"
        Exit Sub

    End If

    If mytablex.RecordCount > 0 Then
        txtserver = Trim("" & mytablex.Fields("txtserver"))
        txtusername = Trim("" & mytablex.Fields("txtusername"))
        txtpassword = Trim("" & mytablex.Fields("txtpassword"))
        txtfromname = Trim("" & mytablex.Fields("txtfromname"))
        txtfromemail = Trim("" & mytablex.Fields("txtfromemail"))
        txtport = Trim("" & mytablex.Fields("txtport"))
        txtselecciona = Trim("" & mytablex.Fields("txtselecciona"))
        chkssl = Trim("" & mytablex.Fields("chkssl"))
        txtto = Trim("" & mytablex.Fields("txtfromemail"))
        txtattach = creafilerar 'Rar Archivo Adjunto
        MsgBox ("ENVIANDO CORREO")
        txtsubject = Trim("" & mytablex.Fields("txtsubject"))
        txtmsg = Trim("" & mytablex.Fields("txtmsg"))
        txtmsg = txtmsg & Chr$(10) & Chr$(13) & ""
        txtmsg = txtmsg & Format(Now, "dd/mm/yyyy") + " " + Format(Now, "hh:mm:ss")

        If Len(Trim("" & mytablex.Fields("txtfromemail"))) > 0 Then
            txtto = Trim("" & mytablex.Fields("txtfromemail"))
            retval = SendMail(Trim$(txtto), Trim$(txtsubject), Trim$(txtfromname) & "<" & Trim$(txtfromemail) & ">", Trim$(txtmsg), Trim$(txtserver), CInt(Trim$(txtport)), Trim$(txtusername), Trim$(txtpassword), Trim$(txtattach), True, txtselecciona, txthtml)
   
        End If

        MsgBox "Correo Enviado ", 48, "Aviso"

    End If

    mytablex.Close

    Exit Sub
cmd0905677_err:
    MsgBox "No se Pudo enviar Correo... " + error$, 48, "Aviso"
    Exit Sub

End Sub

