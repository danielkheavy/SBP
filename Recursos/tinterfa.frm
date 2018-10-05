VERSION 5.00
Begin VB.Form tinterfa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Interfase Tablas de Producto Sistema Orion"
   ClientHeight    =   3180
   ClientLeft      =   150
   ClientTop       =   750
   ClientWidth     =   8820
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox list1 
      Height          =   315
      Left            =   5400
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox caja 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   10
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox almacen 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   8
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox fechaf 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   5
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox fechai 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox procesar 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label qap 
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   600
      Width           =   615
   End
   Begin VB.Label dd 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Caja"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bodega"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaFinal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaInicio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label xfecha 
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Digite Un Comando:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Menu ldfso2323 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tinterfa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim orionv5 As String

Private Sub Command1_Click()

    If procesar = "MIAS" Then
        CLIENTESMIA

        'pasa_productos
    End If

    If procesar = "ASIA" Then
        precios_asia
        Exit Sub

    End If

    If procesar = "LARITZA" Then
        graba_familialaritza
        'graba_productolaritza

    End If

    If procesar = "PERCEPCION" Then
        graba_percepcion

    End If

    If procesar = "CLIENTES" Then
        graba_percepcion_clientes

    End If

    If procesar = "RECAVE" Then
        graba_recave
        graba_recave_equiva
        graba_recave_precios
        Exit Sub

    End If

    If procesar = "DAVID" Then
        graba_david
        Exit Sub

    End If

    If procesar = "MONICA" Then
        monica_cuenta
        Exit Sub

    End If

    If procesar = "SISCONT" Then
        siscont_cuenta
        Exit Sub

    End If

    If procesar = "DENISSE" Then
        graba_denisse
        Exit Sub

    End If

    If procesar = "V51" Then
        orionv5 = "D:\rp_orion.v2\001d\06"
        pone_v5almacen
        Exit Sub
        graba_cproformv5
        graba_dproformv5
        carga_paramv5
        Exit Sub
   
        graba_facturav5
        graba_detallev5
        graba_fpagovv5
      
    End If

    If procesar = "V5" Then
        'GoTo POSO
        orionv5 = "d:\rp_orion.v2\001d\06"
        'GoTo POSO
        'graba_almacenv5
        'Exit Sub
        tipodoc_v5
        'Exit Sub
        'MsgBox "abcde"
        'pone_v5precios
        'Exit Sub
   
        graba_v5producto
        pone_v5precios
        'Exit Sub
   
        graba_equivav5
        'Exit Sub
        'orionv5 = "Z:\rp_orion.v2\001d\06"
        proveedor_v5
        pone_v5subfamilia
        'Exit Sub
        'graba_equivav5
        'Exit Sub
        carga_paramv5
        'Exit Sub
        pone_v5saldoini
        'Exit Sub
        pone_v5bodega
        'POSO:
        pone_v5fpago
        'pone_v5precios
        'graba_v5producto
        pone_v5familia
        'pone_v5subfamilia
        pone_v5marca
        'tipodoc_v5
        vendedor_v5
        clientes_v5
   
        parameca_v5
        cuentacd_v5
        recibos_v5
        'Exit Sub
        'graba_facturav5
        'graba_detallev5
        'graba_fpagovv5
        recibos_v5
        cuentac_v5
        cuentacd_v5
   
        Exit Sub

    End If

    If procesar = "RX" Then
        pasar_datauno

    End If

    If procesar = "MARCAS" Then
        graba_xxfamiliasx
        Exit Sub

    End If

    If procesar = "MAXIMO" Then
        MsgBox "Hola"
        orionv4 = "\RP_ORION.V2\resinas"
        almacen = "01"
        ngraba_producto
        Exit Sub

    End If

    If procesar = "CAJAMARCA" Then
        orionv4 = "\cajamarca"
        almacen = "01"
        cajamarca_producto
        Exit Sub

    End If

    If procesar = "Chiclayo" Then
        graba_producto_chiclayo

    End If

    If procesar = "Guido" Then
        orionv4 = "\rp_orion.v2\guido"
        graba_producto_guido

    End If

    If procesar = "CAJAMARCA" Then
        importa_cajamarca

    End If

    If procesar = "Z" Then
        temporal_graba
        Exit Sub

    End If

    If procesar = "X" Then
        orionv4 = "C:\ORION.V4\001D\01"
        cuentas_corrientes

    End If

    If procesar = "XX" Then
        cuentas_corrientes1

    End If

    If procesar = "Importar" Then
        proceso_importar

    End If

    If procesar = "Dona" Then
        graba_dona
        Exit Sub

    End If

    If procesar = "Procesar" Then
        'If Len(almacen) = 0 Then
        '   almacen.SetFocus
        '   Exit Sub
        'End If
        orionv4 = "\ORION.V4\001D\01"
        almacen = "01"
        'graba_almacenv4
        '   Exit Sub'
        'graba_receta
        'Exit Sub
        graba_producto
        'Exit Sub
        'Exit Sub
        'graba_dona
        graba_clientes
        graba_proveedor
        'Exit Sub
        graba_equiva
        graba_familia
        graba_subfamilia
        graba_categoria
        graba_color
        graba_marca
        graba_seccion
        'MsgBox "abc"
        graba_proveedor
        graba_linea

    End If

    If procesar = "ORIONV2" Then
        'If Len(almacen) = 0 Then
        '   almacen.SetFocus
        '   Exit Sub
        'End If
        orionv4 = "C:\ORION.V2\001D"
        almacen = "01"
        'graba_almacenv4
        '   Exit Sub'
        'graba_receta
        'Exit Sub
        graba_producto
        graba_familia
        graba_subfamilia
        Exit Sub
        'Exit Sub
        'graba_dona
        graba_clientes
        graba_proveedor
        'Exit Sub
        graba_equiva
        graba_familia
        graba_subfamilia
        graba_categoria
        graba_color
        graba_marca
        graba_seccion
        MsgBox "abc"
        graba_proveedor
        graba_linea

    End If

    If procesar = "Pos" Then
        If Not IsDate(fechai) Then
            fechai.SetFocus
            Exit Sub

        End If

        If Not IsDate(fechaf) Then
            fechaf.SetFocus
            Exit Sub

        End If

        'MsgBox "Hola"
        orionv4 = "Z:\orion.v4\001d\01\2002"
        almacen = "01"
        caja = "01"
        'graba_pedido
        'graba_detalle_pedido
        'graba_ocompra
        'graba_detalle_ocompra
   
        graba_ventas
        graba_detalle_ventas
        graba_fpago_ventas
        'graba_ingresos
        'cuentas_corrientes
        'pafacre
        MsgBox "Abace"

        'cuentas_corrientes1
    End If

    If procesar = "Exporta" Then
        If Not IsDate(fechai) Then
            fechai.SetFocus
            Exit Sub

        End If

        If Not IsDate(fechaf) Then
            fechaf.SetFocus
            Exit Sub

        End If

        'MsgBox "Hola"
        graba_ventas1
        graba_detalle_ventas1

        'graba_fpago_ventas1
    End If

End Sub

Sub graba_subfamilia()

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    cn.Execute ("delete from subfamil")
    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("subfamilia")

    Do

        If mytabley.EOF Then Exit Do
        mytablex.Open "select * from subfamil where familia='" & "" & mytabley.Fields("familia") & "' and subfamilia='" & "" & mytabley.Fields("subfamilia") & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.AddNew
            mytablex.Fields("familia") = "" & mytabley.Fields("familia")
            mytablex.Fields("subfamilia") = "" & mytabley.Fields("subfamilia")
            mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcio")
            mytablex.Update
        Else
   
            mytablex.Fields("familia") = "" & mytabley.Fields("familia")
            mytablex.Fields("subfamilia") = "" & mytabley.Fields("subfamilia")
            mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcio")
            mytablex.Update

        End If

        mytablex.Close
        mytabley.MoveNext
    Loop
    '------------------------------------- ------------

    MsgBox "Subfamil proceso Terminado", 48, "Aviso"

End Sub

Sub graba_familia()

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    cn.Execute ("delete from familia")
    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("familia")

    Do

        If mytabley.EOF Then Exit Do
        mytablex.Open "select * from familia where familia='" & "" & mytabley.Fields("familia") & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.AddNew
            mytablex.Fields("familia") = "" & mytabley.Fields("familia")
            mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcio")
            mytablex.Fields("vetouch") = "S"
            mytablex.Update
        Else
            mytablex.Fields("familia") = "" & mytabley.Fields("familia")
            mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcio")
            mytablex.Fields("vetouch") = "S"
            mytablex.Update

        End If

        mytablex.Close
        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    MsgBox "Familia proceso Terminado", 48, "Aviso"

End Sub

Sub graba_producto()

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset  'productos

    Dim mytabley As Table

    Dim vr

    Dim sdx      As Double

    Dim mytableb As Table 'almacen

    Dim mytablec As Table 'almacen orion ant

    sdx = 0
    cn.Execute ("delete from producto")
    'cn.Execute ("delete from precios")
    'cn.Execute ("delete from dueno")
    'cn.Execute ("delete from codprov")

    'MsgBox orionv4
    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("producto")

    Do

        If mytabley.EOF Then Exit Do
        mytablex.Open "select * from producto where producto='" & "" & mytabley.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.AddNew
            pone_registro mytablex, mytabley
            mytablex.Update
        Else
            pone_registro mytablex, mytabley
            mytablex.Update

        End If

        mytablex.Close
        sdx = sdx + 1
        vr = DoEvents()
        dd = "" & sdx
        mytabley.MoveNext
    Loop
    mytabley.Close
    'mytableb.Close
    'mytablec.Close
    mydby.Close
    MsgBox "Producto proceso Terminado", 48, "Aviso"

End Sub

Sub pone_registro(mytablex As ADODB.Recordset, mytabley As Table)

    Dim mytablea As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    mytablex.Fields("isc") = Val("" & mytabley.Fields("ISC"))
    'mytablex.Fields("ivap") = Val("" & mytabley.Fields("nodscto")) 'ivap
    'Exit Sub

    mytablex.Fields("producto") = Trim("" & mytabley.Fields("producto"))
    mytablex.Fields("barras") = "" & mytabley.Fields("barras")
    mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcio")
    mytablex.Fields("descorto") = "" & mytabley.Fields("abreviado")

    mytablex.Fields("presenta") = "" & mytabley.Fields("presentaci")
    mytablex.Fields("dsctoref") = Val("" & mytabley.Fields("retencion"))

    mytablex.Fields("familia") = "" & mytabley.Fields("familia")
    mytablex.Fields("subfamilia") = "" & mytabley.Fields("subfamilia")
    mytablex.Fields("seccion") = "" & mytabley.Fields("seccion")
    mytablex.Fields("marca") = "" & mytabley.Fields("marca")
    mytablex.Fields("categoria") = "" & mytabley.Fields("categoria")
    'mytablex.Fields("linea") = "" & mytabley.Fields("flagtalla")
    mytablex.Fields("color") = "" & mytabley.Fields("color")
    mytablex.Fields("fabrica") = ""
    mytablex.Fields("serie") = ""
    mytablex.Fields("peso") = "" & mytabley.Fields("balanza")
    mytablex.Fields("servicio") = ""
    mytablex.Fields("vecaja") = "S"
    mytablex.Fields("igv") = Val("" & mytabley.Fields("igv"))
    mytablex.Fields("isc") = 0
    mytablex.Fields("pesokgr") = 0.001
    mytablex.Fields("comision") = 0
    mytablex.Fields("monedac") = "" & mytabley.Fields("monedac")
    mytablex.Fields("unidad") = "" & mytabley.Fields("unidad")
    mytablex.Fields("factor") = Val("" & mytabley.Fields("factor"))
    mytablex.Fields("costou") = Val("" & mytabley.Fields("costopaqu"))
    mytablex.Fields("costop") = Val("" & mytabley.Fields("costopaqp"))
    mytablex.Fields("monedav") = "" & mytabley.Fields("moneda")
    mytablex.Fields("estado") = "S"
    'mytablex.Fields("minimo") = Val("" & mytabley.Fields("stkminimo"))
    'mytablex.Fields("maximo") = Val("" & mytabley.Fields("stkmaximo"))

    'mytablez.Open "select * from dueno where codigo='" & "" & mytabley.Fields("seccion") & "' and local='01' and producto='" & "" & mytabley.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic
    'If mytablez.RecordCount = 0 Then
    '   mytablez.AddNew
    '   mytablez.Fields("codigo") = "" & mytabley.Fields("seccion")
    '   mytablez.Fields("local") = "01"
    '   mytablez.Fields("producto") = "" & mytabley.Fields("producto")
    '   mytablez.Update
    'Else
   
    '   mytablez.Fields("codigo") = "" & mytabley.Fields("seccion")
    '   mytablez.Fields("local") = "01"
    '   mytablez.Fields("producto") = "" & mytabley.Fields("producto")
    '   mytablez.Update
    'End If
    'mytablez.Close
    'mytablez.Open "select * from codprov where codigo='" & "" & mytabley.Fields("proveedor") & "' and producto='" & mytabley.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic
    'If mytablez.RecordCount = 0 Then
    '   mytablez.AddNew
    '   pone_detalle mytablez, mytabley, 0
    '   mytablez.Update
    'Else
   
    '   pone_detalle mytablez, mytabley, 0
    '   mytablez.Update
    'End If
    'mytablez.Close

    'mytablez.Open "select * from codprov where codigo='" & "" & mytabley.Fields("proveedor1") & "' and producto='" & mytabley.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic
    'If mytablez.RecordCount = 0 Then
    '   mytablez.AddNew
    '   pone_detalle mytablez, mytabley, 1
    '   mytablez.Update
    'Else
 
    '   pone_detalle mytablez, mytabley, 1
    '   mytablez.Update
    'End If
    'mytablez.Close

    'grabando precios al local nro 1
    mytablea.Open "select * from precios where producto='" & "" & mytabley.Fields("producto") & "' and local='01'", cn, adOpenStatic, adLockOptimistic

    If mytablea.RecordCount = 0 Then
        mytablea.AddNew
        pone_detalle01 mytablea, mytabley, "01"
        mytablea.Update
    Else
        pone_detalle01 mytablea, mytabley, "01"
        mytablea.Update

    End If

    mytablea.Close
    'mytablea.Open "select * from precios where producto='" & "" & mytabley.Fields("producto") & "' and local='02'", cn, adOpenStatic, adLockOptimistic
    'If mytablea.RecordCount = 0 Then '
    '   mytablea.AddNew
    '   pone_detalle01 mytablea, mytabley, "02"
    '   mytablea.Update
    '   Else
    '   pone_detalle01 mytablea, mytabley, "02"
    '   mytablea.Update
    'End If
    'mytablea.Close
    'mytablea.Open "select * from precios where producto='" & "" & mytabley.Fields("producto") & "' and local='03'", cn, adOpenStatic, adLockOptimistic
    'If mytablea.RecordCount = 0 Then
    '   mytablea.AddNew
    '   pone_detalle01 mytablea, mytabley, "03"
    '   mytablea.Update
    '   Else
    '   pone_detalle01 mytablea, mytabley, "03"
    '   mytablea.Update
    'End If
    'mytablea.Close
    'mytablea.Open "select * from precios where producto='" & "" & mytabley.Fields("producto") & "' and local='04'", cn, adOpenStatic, adLockOptimistic
    'If mytablea.RecordCount = 0 Then
    '   mytablea.AddNew
    '   pone_detalle01 mytablea, mytabley, "04"
    '   mytablea.Update
    '   Else
    '   pone_detalle01 mytablea, mytabley, "04"
    '   mytablea.Update
    'End If
    'mytablea.Close

    'grabando almacen
    'mytablec.Seek "=", "" & mytabley.Fields("producto"), "01"
    'If Not mytablec.NoMatch Then
    'mytableb.Seek "=", "01", "" & mytablec.Fields("producto"), almacen
    'If mytableb.NoMatch Then
    '   mytableb.AddNew
    '   mytableb.Fields("local") = "01"
    '   mytableb.Fields("producto") = "" & mytablec.Fields("producto")
    '   mytableb.Fields("bodega") = almacen
    '   mytableb.Fields("saldo") = Val("" & mytablec.Fields("saldo"))
    '   mytableb.Update
    'End If
    'If Not mytableb.NoMatch Then
    '   mytableb.Edit
    '   mytableb.Fields("local") = "01"
    '   mytableb.Fields("producto") = "" & mytablec.Fields("producto")
    '   mytableb.Fields("bodega") = almacen
    '   mytableb.Fields("saldo") = Val("" & mytablec.Fields("saldo"))
    '   mytableb.Update
    'End If
    'End If
End Sub

Sub pone_detalle01(mytablex As ADODB.Recordset, mytabley As Table, buf As String)
    mytablex.Fields("local") = buf
    mytablex.Fields("producto") = "" & mytabley.Fields("producto")
    mytablex.Fields("ccosto") = "" & mytabley.Fields("seccion")
    'mytablex.Fields("monedav") = "" & mytabley.Fields("moneda")
    mytablex.Fields("factor1") = Val("" & mytabley.Fields("factor1"))
    mytablex.Fields("unidad1") = "" & mytabley.Fields("unidad1")
    mytablex.Fields("pventa1") = Val("" & mytabley.Fields("pventa1"))

    mytablex.Fields("factor2") = Val("" & mytabley.Fields("factor2"))
    mytablex.Fields("unidad2") = "" & mytabley.Fields("unidad2")
    mytablex.Fields("pventa2") = Val("" & mytabley.Fields("pventa2"))

    mytablex.Fields("factor3") = Val("" & mytabley.Fields("factor3"))
    mytablex.Fields("unidad3") = "" & mytabley.Fields("unidad3")
    mytablex.Fields("pventa3") = Val("" & mytabley.Fields("pventa3"))

    mytablex.Fields("factor4") = Val("" & mytabley.Fields("factor4"))
    mytablex.Fields("unidad4") = "" & mytabley.Fields("unidad4")
    mytablex.Fields("pventa4") = Val("" & mytabley.Fields("pventa4"))

    mytablex.Fields("factor5") = Val("" & mytabley.Fields("factor5"))
    mytablex.Fields("unidad5") = "" & mytabley.Fields("unidad5")
    mytablex.Fields("pventa5") = Val("" & mytabley.Fields("pventa5"))

    mytablex.Fields("factor6") = Val("" & mytabley.Fields("factor6"))
    mytablex.Fields("unidad6") = "" & mytabley.Fields("unidad6")
    mytablex.Fields("pventa6") = Val("" & mytabley.Fields("pventa6"))

    mytablex.Fields("factor7") = Val("" & mytabley.Fields("factor7"))
    mytablex.Fields("unidad7") = "" & mytabley.Fields("unidad7")
    mytablex.Fields("pventa7") = Val("" & mytabley.Fields("pventa7"))

    mytablex.Fields("factor8") = Val("" & mytabley.Fields("factor8"))
    mytablex.Fields("unidad8") = "" & mytabley.Fields("unidad8")
    mytablex.Fields("pventa8") = Val("" & mytabley.Fields("pventa8"))

    mytablex.Fields("factor9") = Val("" & mytabley.Fields("factor9"))
    mytablex.Fields("unidad9") = "" & mytabley.Fields("unidad9")
    mytablex.Fields("pventa9") = Val("" & mytabley.Fields("pventa9"))

    mytablex.Fields("factor10") = Val("" & mytabley.Fields("factor10"))
    mytablex.Fields("unidad10") = "" & mytabley.Fields("unidad10")
    mytablex.Fields("pventa10") = Val("" & mytabley.Fields("pventa10"))

    mytablex.Fields("minimo11") = Val("" & mytabley.Fields("p1min"))
    mytablex.Fields("minimo12") = Val("" & mytabley.Fields("p2min"))
    mytablex.Fields("minimo13") = Val("" & mytabley.Fields("p3min"))
    mytablex.Fields("minimo14") = Val("" & mytabley.Fields("p4min"))

    mytablex.Fields("maximo11") = Val("" & mytabley.Fields("p1max"))
    mytablex.Fields("maximo12") = Val("" & mytabley.Fields("p2max"))
    mytablex.Fields("maximo13") = Val("" & mytabley.Fields("p3max"))
    mytablex.Fields("maximo14") = Val("" & mytabley.Fields("p4max"))

    mytablex.Fields("pventa11") = Val("" & mytabley.Fields("ppventa1"))
    mytablex.Fields("pventa12") = Val("" & mytabley.Fields("ppventa2"))
    mytablex.Fields("pventa13") = Val("" & mytabley.Fields("ppventa4"))
    mytablex.Fields("pventa14") = Val("" & mytabley.Fields("ppventa5"))

End Sub

Sub pone_detalle(mytablex As ADODB.Recordset, rs As Table, sw As Integer)
    mytablex.Fields("producto") = "" & rs.Fields("producto")

    If sw = 0 Then
        mytablex.Fields("codigo") = "" & rs.Fields("proveedor")
        mytablex.Fields("codigop") = "" & rs.Fields("cpa")

    End If

    If sw = 1 Then
        mytablex.Fields("codigo") = "" & rs.Fields("proveedor1")
        mytablex.Fields("codigop") = "" & rs.Fields("cpa1")

    End If

End Sub

Private Sub Form_Load()
    fechai = Format(Now, "dd/mm/yyyy")
    fechaf = Format(Now, "dd/mm/yyyy")

End Sub

Private Sub ldfso2323_Click()
    tinterfa.Hide
    Unload tinterfa

End Sub

Sub graba_categoria()

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    If procesar <> "Procesar" Then
        procesar = ""
        procesar.SetFocus
        Exit Sub

    End If

    cn.Execute ("delete from categori")
    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("categori")
    Do

        If mytabley.EOF Then Exit Do
        mytablex.Open "select * from categori where categoria='" & "" & mytabley.Fields("categoria") & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.AddNew
            mytablex.Fields("categoria") = "" & mytabley.Fields("categoria")
            mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcio")
            mytablex.Update
        Else
            mytablex.Fields("categoria") = "" & mytabley.Fields("categoria")
            mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcio")
            mytablex.Update

        End If

        mytablex.Close
        mytabley.MoveNext
    Loop

    MsgBox "Categoria proceso Terminado", 48, "Aviso"

End Sub

Sub graba_color()

    Dim mydby    As Database

    Dim mytablex As Table

    Dim mytabley As Table

    If procesar <> "Procesar" Then
        procesar = ""
        procesar.SetFocus
        Exit Sub

    End If

    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("color")

    Set mytablex = mydbxglo.OpenTable("color")
    mytablex.Index = "color"
    Do

        If mytabley.EOF Then Exit Do
        mytablex.Seek "=", "" & mytabley.Fields("color")

        If mytablex.NoMatch Then
            mytablex.AddNew
            mytablex.Fields("color") = "" & mytabley.Fields("color")
            mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcio")
            mytablex.Update

        End If

        If Not mytablex.NoMatch Then
            mytablex.Edit
            mytablex.Fields("color") = "" & mytabley.Fields("color")
            mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcio")
            mytablex.Update

        End If

        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    mytablex.Close
    MsgBox "Color proceso Terminado", 48, "Aviso"

End Sub

Sub graba_marca()

    Dim mydby As Database

    Dim vr

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    If procesar <> "Procesar" Then
        procesar = ""
        procesar.SetFocus
        Exit Sub

    End If

    cn.Execute ("delete from marca")

    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("marca")

    'list1.Clear
    'MsgBox "marca"
    Do

        If mytabley.EOF Then Exit Do
        'mytablex.Open "select * from marca where marca='" & "" & mytabley.Fields("marca") & "'", cn, adOpenStatic, adLockOptimistic
        mytablex.Open "select * from marca where marca='" & Trim("" & mytabley.Fields("marca")) & "'", cn, adOpenStatic, adLockOptimistic

        'dd = "" & mytabley.Fields("marca")
        'vr = DoEvents
        'list1.AddItem "" & mytabley.Fields("marca")
        If mytablex.RecordCount = 0 Then
            mytablex.AddNew
            mytablex.Fields("marca") = Trim("" & mytabley.Fields("marca"))
            mytablex.Fields("descripcio") = Trim("" & mytabley.Fields("marca"))
            mytablex.Update

            'Else
            '   mytablex.Fields("marca") = Trim("" & mytabley.Fields("marca"))
            '   mytablex.Fields("descripcio") = Trim("" & mytabley.Fields("descripcio"))
            '   mytablex.Update
        End If

        mytablex.Close
        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    MsgBox "Marca proceso Terminado", 48, "Aviso"

End Sub

Sub graba_talla()

    Dim mydby    As Database

    Dim mytablex As Table

    Dim mytabley As Table

    If procesar <> "Procesar" Then
        procesar = ""
        procesar.SetFocus
        Exit Sub

    End If

    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("familia")

    Set mytablex = mydbxglo.OpenTable("familia")
    mytablex.Index = "familia"
    Do

        If mytabley.EOF Then Exit Do
        mytablex.Seek "=", "" & mytabley.Fields("familia")

        If mytablex.NoMatch Then
            mytablex.AddNew
            mytablex.Fields("familia") = "" & mytabley.Fields("familia")
            mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcio")
            mytablex.Update

        End If

        If Not mytablex.NoMatch Then
            mytablex.Edit
            mytablex.Fields("familia") = "" & mytabley.Fields("familia")
            mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcio")
            mytablex.Update

        End If

        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    mytablex.Close
    MsgBox "proceso Terminado", 48, "Aviso"

End Sub

Sub graba_linea()

    Dim mydby    As Database

    Dim mytablex As Table

    Dim mytabley As Table

    If procesar <> "Procesar" Then
        procesar = ""
        procesar.SetFocus
        Exit Sub

    End If

    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("linea")

    Set mytablex = mydbxglo.OpenTable("linea")
    mytablex.Index = "linea"
    Do

        If mytabley.EOF Then Exit Do
        mytablex.Seek "=", "" & mytabley.Fields("linea")

        If mytablex.NoMatch Then
            mytablex.AddNew
            mytablex.Fields("linea") = "" & mytabley.Fields("linea")
            mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcio")
            mytablex.Fields("t1") = "" & mytabley.Fields("t1")
            mytablex.Fields("t2") = "" & mytabley.Fields("t2")
            mytablex.Fields("t3") = "" & mytabley.Fields("t3")
            mytablex.Fields("t4") = "" & mytabley.Fields("t4")
            mytablex.Fields("t5") = "" & mytabley.Fields("t5")
            mytablex.Fields("t6") = "" & mytabley.Fields("t6")
            mytablex.Fields("t7") = "" & mytabley.Fields("t7")
            mytablex.Fields("t8") = "" & mytabley.Fields("t8")
            mytablex.Fields("t9") = "" & mytabley.Fields("t9")
            mytablex.Fields("t10") = "" & mytabley.Fields("t10")
            mytablex.Fields("t11") = "" & mytabley.Fields("t11")
            mytablex.Fields("t12") = "" & mytabley.Fields("t12")
            mytablex.Fields("t13") = "" & mytabley.Fields("t13")
            mytablex.Fields("t14") = "" & mytabley.Fields("t14")
            mytablex.Fields("t15") = "" & mytabley.Fields("t15")
            mytablex.Fields("t16") = "" & mytabley.Fields("t16")
            mytablex.Update

        End If

        If Not mytablex.NoMatch Then
            mytablex.Edit
            mytablex.Fields("linea") = "" & mytabley.Fields("linea")
            mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcio")
            mytablex.Fields("t1") = "" & mytabley.Fields("t1")
            mytablex.Fields("t2") = "" & mytabley.Fields("t2")
            mytablex.Fields("t3") = "" & mytabley.Fields("t3")
            mytablex.Fields("t4") = "" & mytabley.Fields("t4")
            mytablex.Fields("t5") = "" & mytabley.Fields("t5")
            mytablex.Fields("t6") = "" & mytabley.Fields("t6")
            mytablex.Fields("t7") = "" & mytabley.Fields("t7")
            mytablex.Fields("t8") = "" & mytabley.Fields("t8")
            mytablex.Fields("t9") = "" & mytabley.Fields("t9")
            mytablex.Fields("t10") = "" & mytabley.Fields("t10")
            mytablex.Fields("t11") = "" & mytabley.Fields("t11")
            mytablex.Fields("t12") = "" & mytabley.Fields("t12")
            mytablex.Fields("t13") = "" & mytabley.Fields("t13")
            mytablex.Fields("t14") = "" & mytabley.Fields("t14")
            mytablex.Fields("t15") = "" & mytabley.Fields("t15")
            mytablex.Fields("t16") = "" & mytabley.Fields("t16")
            mytablex.Update

        End If

        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    mytablex.Close
 
    MsgBox "Linea proceso Terminado", 48, "Aviso"

End Sub

Sub graba_seccion()

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    If procesar <> "Procesar" Then
        procesar = ""
        procesar.SetFocus
        Exit Sub

    End If

    cn.Execute ("delete from seccion")
    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("Seccion")

    Do

        If mytabley.EOF Then Exit Do
        mytablex.Open "select * from seccion where seccion='" & "" & mytabley.Fields("seccion") & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.AddNew
            mytablex.Fields("Seccion") = "" & mytabley.Fields("Seccion")
            mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcio")
            mytablex.Update
        Else
            mytablex.Fields("Seccion") = "" & mytabley.Fields("Seccion")
            mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcio")
            mytablex.Update

        End If

        mytablex.Close
        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
 
    MsgBox "Seccion proceso Terminado", 48, "Aviso"

End Sub

Sub graba_proveedor()

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    If procesar <> "Procesar" Then
        procesar = ""
        procesar.SetFocus
        Exit Sub

    End If

    'cn.Execute ("delete from proveedo")
    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("proveedo")
    Do

        If mytabley.EOF Then Exit Do
        mytablex.Open "select * from proveedo where codigo='" & "" & mytabley.Fields("codigo") & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.AddNew
            mytablex.Fields("codigo") = "" & mytabley.Fields("codigo")
            mytablex.Fields("codigo1") = "" & mytabley.Fields("ruc")
            mytablex.Fields("nombre") = "" & mytabley.Fields("nombre")
            mytablex.Update
        Else
            mytablex.Fields("codigo") = "" & mytabley.Fields("codigo")
            mytablex.Fields("codigo1") = "" & mytabley.Fields("ruc")
            mytablex.Fields("nombre") = "" & mytabley.Fields("nombre")
            mytablex.Update

        End If

        mytablex.Close
        mytabley.MoveNext
    Loop
    '------------------------------------- ------------

    MsgBox "Proveedor proceso Terminado", 48, "Aviso"

End Sub

Sub graba_clientes()

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    Dim vr

    If procesar <> "Procesar" Then
        procesar = ""
        procesar.SetFocus
        Exit Sub

    End If

    cn.Execute ("delete from clientes")
    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("clientes")
    xfecha = ""

    Do

        If mytabley.EOF Then Exit Do
        If Trim("" & mytabley.Fields("codigo")) > 0 Then
            If mytablex.State = 1 Then mytablex.Close
            mytablex.Open "select * from clientes where codigo='" & Trim("" & mytabley.Fields("codigo")) & "'", cn, adOpenStatic, adLockOptimistic

            If mytablex.RecordCount = 0 Then
                mytablex.AddNew
                mytablex.Fields("codigo") = Trim("" & mytabley.Fields("codigo"))
                mytablex.Fields("nombre") = Trim("" & mytabley.Fields("nombre"))
                mytablex.Fields("direccion") = Trim("" & mytabley.Fields("direccion"))
                mytablex.Fields("descuento") = Val("" & mytabley.Fields("descuento"))
                mytablex.Update
                vr = DoEvents
                xfecha = "" & mytabley.Fields("nombre")
            Else

                'mytablex.Fields("codigo") = Trim("" & mytabley.Fields("codigo"))
                'mytablex.Fields("nombre") = "" & mytabley.Fields("nombre")
                'mytablex.Update
            End If

        End If

        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    MsgBox "Clientes proceso Terminado", 48, "Aviso"

End Sub

Sub graba_detalle_ventas1()

    Dim mydby    As Database

    Dim mytablex As Table

    Dim mytabley As Snapshot

    Dim buf      As String

    Dim buf1     As String

    Dim vr

    '----eliminando---------
    Set mydby = OpenDatabase(globaldir & "\v4", False, False, "foxpro 2.5;")
    buf = " fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
    mydby.Execute "DELETE FROM detalle where  " & buf
    xfecha = ""
    buf = "select * from detalle where "
    buf = buf & "  fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
    buf = buf & " and (tipo='1' or tipo='2' or tipo='3' or tipo='4') "
    Set mytabley = mydbxglo.CreateSnapshot(buf)
    Set mytablex = mydby.OpenTable("detalle")
    mytablex.Index = "detalle"
    Do

        If mytabley.EOF Then Exit Do
        buf1 = mytabley.Fields("numero")

        If Mid$("" & mytabley.Fields("numero"), 1, 2) <> "B0" And Mid$("" & mytabley.Fields("numero"), 1, 2) <> "F0" Then
            If "" & mytabley.Fields("tipo") = "1" Then
                buf1 = "B" & mytabley.Fields("caja") & "-" & mytabley.Fields("numero")

            End If

            If "" & mytabley.Fields("tipo") = "2" Then
                buf1 = "F" & mytabley.Fields("caja") & "-" & mytabley.Fields("numero")

            End If

        End If

        If Mid$("" & mytabley.Fields("numero"), 4, 1) <> "-" Then
            If "" & mytabley.Fields("tipo") = "3" Then
                buf1 = "" & mytabley.Fields("serie") & "-" & mytabley.Fields("numero")

            End If

            If "" & mytabley.Fields("tipo") = "4" Then
                buf1 = "" & mytabley.Fields("serie") & "-" & mytabley.Fields("numero")

            End If

            If Len(buf1) > 12 Then
                buf1 = Mid$(buf1, 1, 12)

            End If

        End If

        'mytablex.Seek "=", "" & mytabley.Fields("tipo"), buf1
        'If mytablex.NoMatch Then
        mytablex.AddNew
        pone_registro_detalle1 mytablex, mytabley, buf1
        mytablex.Update
        vr = DoEvents
        'End If
        'If Not mytablex.NoMatch Then
        '   mytablex.Edit
        '   pone_registro_detalle1 mytablex, mytabley, buf1
        '   mytablex.Update
        '   vr = DoEvents
        'End If
        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    mytabley.Close
    mytablex.Close
    MsgBox "proceso Terminado", 48, "Aviso"

End Sub

Sub pone_registro_detalle1(mytablez As Table, mytabler As Snapshot, buf1 As String)
    mytablez.Fields("estado") = "" & mytabler.Fields("estado")
    mytablez.Fields("tipo") = "" & mytabler.Fields("tipo")
    mytablez.Fields("serie") = Mid$("" & mytabler.Fields("serie"), 1, 3)
    mytablez.Fields("numero") = buf1 'Mid$("" & mytabler.Fields("numero"), 1, 11)
    mytablez.Fields("acu") = "V"
    mytablez.Fields("codigo") = "" & mytabler.Fields("codigo")
    mytablez.Fields("acu1") = "" & mytabler.Fields("acu1")
    mytablez.Fields("fecha") = "" & mytabler.Fields("fecha")
    mytablez.Fields("moneda") = "" & mytabler.Fields("moneda")
    mytablez.Fields("producto") = "" & mytabler.Fields("producto")
    mytablez.Fields("descripcio") = "" & mytabler.Fields("descripcio")
    mytablez.Fields("unidad") = "" & mytabler.Fields("unidad")
    mytablez.Fields("factor") = Val("" & mytabler.Fields("factor"))
    mytablez.Fields("cantidad") = Val("" & mytabler.Fields("cantidad"))
    mytablez.Fields("precio") = Val("" & mytabler.Fields("precio"))
    mytablez.Fields("igv") = Val("" & mytabler.Fields("igv"))
    mytablez.Fields("bruto") = Val("" & mytabler.Fields("neto"))
    mytablez.Fields("descuento") = Val("" & mytabler.Fields("descuento"))
    mytablez.Fields("subtotal") = Val("" & mytabler.Fields("subtotal"))
    mytablez.Fields("impuesto") = Val("" & mytabler.Fields("impuesto"))
    mytablez.Fields("total") = Val("" & mytabler.Fields("total"))
    mytablez.Fields("estado") = "" & mytabler.Fields("estado")
    mytablez.Fields("fechav") = "" & mytabler.Fields("fecha")
    mytablez.Fields("hora") = "" & mytabler.Fields("hora")
    mytablez.Fields("vendedor") = "" & mytabler.Fields("vendedor")
    mytablez.Fields("bodega") = "" & mytabler.Fields("bodega")
    mytablez.Fields("bodegaf") = ""
    mytablez.Fields("deslipo") = Val("" & mytabler.Fields("deslipo"))
    mytablez.Fields("usuario") = "" & mytabler.Fields("usuario")
    mytablez.Fields("caja") = "" & mytabler.Fields("caja")
    mytablez.Fields("turno") = "" & mytabler.Fields("turno")
    mytablez.Fields("servicio") = "" & mytabler.Fields("servicio")
    mytablez.Fields("local") = "" & mytabler.Fields("local")

End Sub

Sub graba_ventas1()

    Dim mydby    As Database

    Dim mytablex As Table

    Dim mytabley As Snapshot

    Dim buf      As String

    Dim buf1     As String

    Dim vr

    '----eliminando---------
    Set mydby = OpenDatabase(globaldir & "\v4", False, False, "foxpro 2.5;")
    buf = " fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
    mydby.Execute "DELETE FROM cabeza where  " & buf
    xfecha = ""
    buf = "select * from factura where "
    buf = buf & "  fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
    buf = buf & " and (tipo='1' or tipo='2' or tipo='3' or tipo='4') "
    Set mytabley = mydbxglo.CreateSnapshot(buf)
    Set mytablex = mydby.OpenTable("cabeza")
    mytablex.Index = "cabeza"
    Do

        If mytabley.EOF Then Exit Do
        buf1 = mytabley.Fields("numero")

        If Mid$("" & mytabley.Fields("numero"), 1, 2) <> "B0" And Mid$("" & mytabley.Fields("numero"), 1, 2) <> "F0" Then
            If "" & mytabley.Fields("tipo") = "1" Then
                buf1 = "B" & mytabley.Fields("caja") & "-" & mytabley.Fields("numero")

            End If

            If "" & mytabley.Fields("tipo") = "2" Then
                buf1 = "F" & mytabley.Fields("caja") & "-" & mytabley.Fields("numero")

            End If

        End If

        If Mid$("" & mytabley.Fields("numero"), 4, 1) <> "-" Then
            If "" & mytabley.Fields("tipo") = "3" Then
                buf1 = "" & mytabley.Fields("serie") & "-" & mytabley.Fields("numero")

            End If

            If "" & mytabley.Fields("tipo") = "4" Then
                buf1 = "" & mytabley.Fields("serie") & "-" & mytabley.Fields("numero")

            End If

            If Len(buf1) > 12 Then
                buf1 = Mid$(buf1, 1, 12)

            End If

        End If

        mytablex.Seek "=", "" & mytabley.Fields("tipo"), buf1

        If mytablex.NoMatch Then
            mytablex.AddNew
            pone_registro_venta1 mytablex, mytabley, buf1
            mytablex.Update
            vr = DoEvents

        End If

        If Not mytablex.NoMatch Then
            mytablex.Edit
            pone_registro_venta1 mytablex, mytabley, buf1
            mytablex.Update
            vr = DoEvents

        End If

        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    mytabley.Close
    mytablex.Close
    MsgBox "proceso Terminado", 48, "Aviso"

End Sub

Sub graba_ventas()

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Snapshot

    Dim buf      As String

    Dim buf1     As String

    Dim vr

    '----eliminando---------
    'buf = " fecha>=" & "DateValue('" & fechai & "'" & ")"
    'buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
    'buf = buf & " and serie='" & caja & "'"
    'buf = buf & " and bodega='" & almacen & "'"
    'cn.Execute "DELETE FROM factura  where  " & buf
    'cn.Execute "DELETE FROM factura"
    xfecha = ""
    buf = "select * from CADIARIO WHERE "
    buf = buf & " fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
    buf = buf & " and   acu='V' "

    Set mydby = OpenDatabase(orionv4 & "\", False, False, "foxpro 2.5;")
    Set mytabley = mydby.CreateSnapshot(buf)
    mytablex.Open "select * from factura where local like '%'", cn, adOpenStatic, adLockOptimistic
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew
        pone_registro_venta mytablex, mytabley
        mytablex.Update
        vr = DoEvents
        xfecha = "" & mytabley.Fields("fecha")
        mytabley.MoveNext
    Loop
    mytabley.Close
    mytablex.Close
    MsgBox "Factura proceso Terminado", 48, "Aviso"

End Sub

Sub graba_pedido()

    Dim mydby    As Database

    Dim mytablex As Table

    Dim mytabley As Snapshot

    Dim buf      As String

    Dim buf1     As String

    Dim vr

    '----eliminando---------
    buf = " fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
    mydbxglo.Execute "DELETE FROM cpedidov where  " & buf
    xfecha = ""
    buf = "select * from cpedidov where "
    buf = buf & "  fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.CreateSnapshot(buf)
    Set mytablex = mydbxglo.OpenTable("cpedidov")
    mytablex.Index = "tfactura"
    Do

        If mytabley.EOF Then Exit Do
        mytablex.Seek "=", "" & mytabley.Fields("local"), "" & mytabley.Fields("tipo"), "" & mytabley.Fields("serie"), Mid("" & mytabley.Fields("numero"), 1, 11)

        If mytablex.NoMatch Then
            mytablex.AddNew
            pone_registro_pedido mytablex, mytabley
            mytablex.Update
            vr = DoEvents
            xfecha = "" & mytabley.Fields("fecha")

        End If

        If Not mytablex.NoMatch Then
            mytablex.Edit
            pone_registro_pedido mytablex, mytabley
            mytablex.Update
            vr = DoEvents
            xfecha = "" & mytabley.Fields("fecha")

        End If

        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    mytabley.Close
    mytablex.Close
    MsgBox "proceso Terminado", 48, "Aviso"

End Sub

Sub graba_ocompra()

    Dim mydby    As Database

    Dim mytablex As Table

    Dim mytabley As Snapshot

    Dim buf      As String

    Dim buf1     As String

    Dim vr

    '----eliminando---------
    buf = " fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
    mydbxglo.Execute "DELETE FROM cordenc where  " & buf
    xfecha = ""
    buf = "select * from cocompra where "
    buf = buf & "  fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.CreateSnapshot(buf)
    Set mytablex = mydbxglo.OpenTable("cordenc")
    mytablex.Index = "tfactura"
    Do

        If mytabley.EOF Then Exit Do
        mytablex.Seek "=", "" & mytabley.Fields("local"), "" & mytabley.Fields("tipo"), "" & mytabley.Fields("serie"), Mid("" & mytabley.Fields("numero"), 1, 11)

        If mytablex.NoMatch Then
            mytablex.AddNew
            pone_registro_ocompra mytablex, mytabley
            mytablex.Update
            vr = DoEvents
            xfecha = "" & mytabley.Fields("fecha")

        End If

        If Not mytablex.NoMatch Then
            mytablex.Edit
            pone_registro_ocompra mytablex, mytabley
            mytablex.Update
            vr = DoEvents
            xfecha = "" & mytabley.Fields("fecha")

        End If

        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    mytabley.Close
    mytablex.Close
    MsgBox "proceso Terminado", 48, "Aviso"

End Sub

Sub graba_fpago_ventas()

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Snapshot

    Dim buf      As String

    Dim vr

    '--------------eliminando----------------------------------------
    'cn.Execute "DELETE FROM fpagov "
    '----------------------------------------------------------------
    xfecha = ""

    buf = "select * from FPDIARIO WHERE "
    buf = buf & " fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
    buf = buf & " and  ( acu='V') "

    'buf = "select * from fpagov where "
    'buf = buf & "  (acu='V')"

    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.CreateSnapshot(buf)
    mytablex.Open "select * from fpagov where local like '%'", cn, adOpenStatic, adLockOptimistic

    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew
        pone_registro_ventaf mytablex, mytabley
        mytablex.Update
        vr = DoEvents
        xfecha = "" & mytabley.Fields("fecha")
        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    mytabley.Close
    mytablex.Close
    MsgBox "Fpagov-proceso Terminado", 48, "Aviso"

End Sub

Sub graba_detalle_ventas()

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Snapshot

    Dim buf      As String

    Dim vr

    '--------------eliminando----------------------------------------
    'cn.Execute ("delete from detalle  ")
    MsgBox "Buscando"
    buf = "select * from deDIARIO WHERE "
    buf = buf & " fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
    buf = buf & " and  ( acu='V') "

    'buf = "select * from detalle where "
    'buf = buf & "  (acu='V' ) "
    MsgBox "Empiezo...1"

    Set mydby = OpenDatabase(orionv4 & "\", False, False, "foxpro 2.5;")
    Set mytabley = mydby.CreateSnapshot(buf)

    mytablex.Open "select * from detalle ", cn, adOpenStatic, adLockOptimistic
    MsgBox "Ahora si.."
    Do

        If mytabley.EOF Then Exit Do
        xfecha = "" & mytabley.Fields("fecha")
        vr = DoEvents
        mytablex.AddNew
        pone_registro_ventax mytablex, mytabley
        mytablex.Update
        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    mytabley.Close
    mytablex.Close
    MsgBox "Detalle-proceso Terminado", 48, "Aviso"

End Sub

Sub graba_detalle_pedido()

    Dim mydby    As Database

    Dim mytablex As Table

    Dim mytabley As Snapshot

    Dim buf      As String

    Dim vr

    '--------------eliminando----------------------------------------
    buf = " fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
    mydbxglo.Execute "DELETE FROM dpedidov where  " & buf
    '----------------------------------------------------------------
    xfecha = ""
    buf = "select * from dpedidov where "
    buf = buf & "  fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.CreateSnapshot(buf)
    Set mytablex = mydbxglo.OpenTable("dpedidov")
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew
        pone_registro_detalle_pedido mytablex, mytabley
        mytablex.Update
        vr = DoEvents
        xfecha = "" & mytabley.Fields("fecha")
        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    mytabley.Close
    mytablex.Close
    MsgBox "Detalle-proceso Terminado", 48, "Aviso"

End Sub

Sub pone_registro_detalle_pedido(mytablez As Table, mytabler As Snapshot)
    mytablez.Fields("estado") = "" & mytabler.Fields("estado")
    mytablez.Fields("tipo") = "" & mytabler.Fields("tipo")
    mytablez.Fields("serie") = "" & mytabler.Fields("serie")
    mytablez.Fields("numero") = Mid$("" & mytabler.Fields("numero"), 1, 11)
   
    mytablez.Fields("tipoclie") = "C"
    mytablez.Fields("acu") = "I"
   
    mytablez.Fields("codigo") = "" & mytabler.Fields("codigo")
    mytablez.Fields("acu1") = "" & mytabler.Fields("acu1")
    mytablez.Fields("fecha") = "" & mytabler.Fields("fecha")
    mytablez.Fields("moneda") = "" & mytabler.Fields("moneda")
    mytablez.Fields("producto") = "" & mytabler.Fields("producto")
    mytablez.Fields("descripcio") = "" & mytabler.Fields("descripcio")
    mytablez.Fields("unidad") = "" & mytabler.Fields("unidad")
    mytablez.Fields("factor") = Val("" & mytabler.Fields("factor"))
    mytablez.Fields("cantidad") = Val("" & mytabler.Fields("cantidad"))
    mytablez.Fields("precio") = Val("" & mytabler.Fields("precio"))
    mytablez.Fields("igv") = Val("" & mytabler.Fields("igv"))
    mytablez.Fields("neto") = Val("" & mytabler.Fields("bruto"))
    mytablez.Fields("descuento") = Val("" & mytabler.Fields("descuento"))
    mytablez.Fields("subtotal") = Val("" & mytabler.Fields("subtotal"))
    mytablez.Fields("impuesto") = Val("" & mytabler.Fields("impuesto"))
    mytablez.Fields("total") = Val("" & mytabler.Fields("total"))
    mytablez.Fields("estado") = "" & mytabler.Fields("estado")
    mytablez.Fields("fechacrea") = "" & mytabler.Fields("fecha")
    mytablez.Fields("hora") = "" & mytabler.Fields("hora")
    mytablez.Fields("vendedor") = "" & mytabler.Fields("vendedor")
    mytablez.Fields("bodega") = almacen
    mytablez.Fields("bodega") = "" & mytabler.Fields("bodega")
    mytablez.Fields("bodegaf") = ""
    mytablez.Fields("deslipo") = Val("" & mytabler.Fields("deslipo"))
    mytablez.Fields("flage") = ""
    mytablez.Fields("linea") = ""
    mytablez.Fields("t1") = 0
    mytablez.Fields("t2") = 0
    mytablez.Fields("t3") = 0
    mytablez.Fields("t4") = 0
    mytablez.Fields("t5") = 0
    mytablez.Fields("t6") = 0
    mytablez.Fields("t7") = 0
    mytablez.Fields("t8") = 0
    mytablez.Fields("t9") = 0
    mytablez.Fields("t10") = 0
    mytablez.Fields("t11") = 0
    mytablez.Fields("t12") = 0
    mytablez.Fields("t13") = 0
    mytablez.Fields("t14") = 0
    mytablez.Fields("t15") = 0
    mytablez.Fields("t16") = 0
    mytablez.Fields("l1") = ""
    mytablez.Fields("l2") = ""
    mytablez.Fields("l3") = ""
    mytablez.Fields("l4") = ""
    mytablez.Fields("local") = ""
    mytablez.Fields("proveedorp") = ""
    mytablez.Fields("observa1") = ""
    mytablez.Fields("observa2") = ""
    mytablez.Fields("observa3") = ""
    mytablez.Fields("observa4") = ""
    mytablez.Fields("zona") = ""
    mytablez.Fields("isc") = 0
    mytablez.Fields("tax") = 0
    mytablez.Fields("vtaneta") = 0
    mytablez.Fields("tcosto") = 0
    mytablez.Fields("ganancia") = 0
    mytablez.Fields("comision") = 0
    mytablez.Fields("usuario") = "" & mytabler.Fields("usuario")
    mytablez.Fields("caja") = "" & mytabler.Fields("caja")
    mytablez.Fields("turno") = "" & mytabler.Fields("turno")
    mytablez.Fields("servicio") = "" & mytabler.Fields("servicio")
    mytablez.Fields("comanda") = ""
    mytablez.Fields("mesa") = ""
    mytablez.Fields("salon") = ""
    mytablez.Fields("mesero") = ""
    'mytablez.Fields("codigop") = ""
    mytablez.Fields("local") = "" & mytabler.Fields("local")

End Sub

Sub graba_detalle_ocompra()

    Dim mydby    As Database

    Dim mytablex As Table

    Dim mytabley As Snapshot

    Dim buf      As String

    Dim vr

    '--------------eliminando----------------------------------------
    buf = " fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
    mydbxglo.Execute "DELETE FROM dordenc where  " & buf
    '----------------------------------------------------------------
    xfecha = ""
    buf = "select * from docompra where "
    buf = buf & "  fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.CreateSnapshot(buf)
    Set mytablex = mydbxglo.OpenTable("dordenc")
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew
        pone_registro_detalle_ocompra mytablex, mytabley
        mytablex.Update
        vr = DoEvents
        xfecha = "" & mytabley.Fields("fecha")
        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    mytabley.Close
    mytablex.Close
    MsgBox "Detalle-proceso Terminado", 48, "Aviso"

End Sub

Sub pone_registro_detalle_ocompra(mytablez As Table, mytabler As Snapshot)
    mytablez.Fields("estado") = "" & mytabler.Fields("estado")
    mytablez.Fields("tipo") = "" & mytabler.Fields("tipo")
    mytablez.Fields("serie") = "" & mytabler.Fields("serie")
    mytablez.Fields("numero") = Mid$("" & mytabler.Fields("numero"), 1, 11)
   
    mytablez.Fields("tipoclie") = "P"
    mytablez.Fields("acu") = "R"
   
    mytablez.Fields("codigo") = "" & mytabler.Fields("codigo")
    mytablez.Fields("acu1") = "" & mytabler.Fields("acu1")
    mytablez.Fields("fecha") = "" & mytabler.Fields("fecha")
    mytablez.Fields("moneda") = "" & mytabler.Fields("moneda")
    mytablez.Fields("producto") = "" & mytabler.Fields("producto")
    mytablez.Fields("descripcio") = "" & mytabler.Fields("descripcio")
    mytablez.Fields("unidad") = "" & mytabler.Fields("unidad")
    mytablez.Fields("factor") = Val("" & mytabler.Fields("factor"))
    mytablez.Fields("cantidad") = Val("" & mytabler.Fields("cantidad"))
    mytablez.Fields("precio") = Val("" & mytabler.Fields("precio"))
    mytablez.Fields("igv") = Val("" & mytabler.Fields("igv"))
    mytablez.Fields("neto") = Val("" & mytabler.Fields("bruto"))
    mytablez.Fields("descuento") = Val("" & mytabler.Fields("descuento"))
    mytablez.Fields("subtotal") = Val("" & mytabler.Fields("subtotal"))
    mytablez.Fields("impuesto") = Val("" & mytabler.Fields("impuesto"))
    mytablez.Fields("total") = Val("" & mytabler.Fields("total"))
    mytablez.Fields("estado") = "" & mytabler.Fields("estado")
    mytablez.Fields("fechacrea") = "" & mytabler.Fields("fecha")
    mytablez.Fields("hora") = "" & mytabler.Fields("hora")
    mytablez.Fields("vendedor") = "" & mytabler.Fields("vendedor")
    mytablez.Fields("bodega") = almacen
    mytablez.Fields("bodega") = "" & mytabler.Fields("bodega")
    mytablez.Fields("bodegaf") = ""
    mytablez.Fields("deslipo") = Val("" & mytabler.Fields("deslipo"))
    mytablez.Fields("flage") = ""
    mytablez.Fields("linea") = ""
    mytablez.Fields("t1") = 0
    mytablez.Fields("t2") = 0
    mytablez.Fields("t3") = 0
    mytablez.Fields("t4") = 0
    mytablez.Fields("t5") = 0
    mytablez.Fields("t6") = 0
    mytablez.Fields("t7") = 0
    mytablez.Fields("t8") = 0
    mytablez.Fields("t9") = 0
    mytablez.Fields("t10") = 0
    mytablez.Fields("t11") = 0
    mytablez.Fields("t12") = 0
    mytablez.Fields("t13") = 0
    mytablez.Fields("t14") = 0
    mytablez.Fields("t15") = 0
    mytablez.Fields("t16") = 0
    mytablez.Fields("l1") = ""
    mytablez.Fields("l2") = ""
    mytablez.Fields("l3") = ""
    mytablez.Fields("l4") = ""
    mytablez.Fields("local") = ""
    mytablez.Fields("proveedorp") = ""
    mytablez.Fields("observa1") = ""
    mytablez.Fields("observa2") = ""
    mytablez.Fields("observa3") = ""
    mytablez.Fields("observa4") = ""
    mytablez.Fields("zona") = ""
    mytablez.Fields("isc") = 0
    mytablez.Fields("tax") = 0
    mytablez.Fields("vtaneta") = 0
    mytablez.Fields("tcosto") = 0
    mytablez.Fields("ganancia") = 0
    mytablez.Fields("comision") = 0
    mytablez.Fields("usuario") = "" & mytabler.Fields("usuario")
    mytablez.Fields("caja") = "" & mytabler.Fields("caja")
    mytablez.Fields("turno") = "" & mytabler.Fields("turno")
    mytablez.Fields("servicio") = "" & mytabler.Fields("servicio")
    mytablez.Fields("comanda") = ""
    mytablez.Fields("mesa") = ""
    mytablez.Fields("salon") = ""
    mytablez.Fields("mesero") = ""
    'mytablez.Fields("codigop") = ""
    mytablez.Fields("local") = "" & mytabler.Fields("local")

End Sub

Sub pone_registro_ventaf(mytablez As ADODB.Recordset, mytabler As Snapshot)
    mytablez.Fields("estado") = "" & mytabler.Fields("estado")
    mytablez.Fields("tipo") = "" & mytabler.Fields("tipo")
    mytablez.Fields("serie") = "" & mytabler.Fields("serie")
    mytablez.Fields("numero") = Mid$("" & mytabler.Fields("numero"), 1, 11)
   
    If "" & mytabler.Fields("acu") = "V" Then
        mytablez.Fields("tipoclie") = "C"

        If "" & mytabler.Fields("tipo") = "1" Then  'ticket boleta
            mytablez.Fields("serie") = "B" & mytabler.Fields("caja")
            mytablez.Fields("numero") = Mid$("" & mytabler.Fields("numero"), 5, 11)
            mytablez.Fields("acu") = "C"

        End If

        If "" & mytabler.Fields("tipo") = "2" Then  'ticket factura
            mytablez.Fields("serie") = "F" & mytabler.Fields("caja")
            mytablez.Fields("numero") = Mid$("" & mytabler.Fields("numero"), 5, 11)
            mytablez.Fields("acu") = "D"

        End If

        If "" & mytabler.Fields("tipo") = "3" Then  'boleta
            mytablez.Fields("numero") = Mid$("" & mytabler.Fields("numero"), 5, 11)
            mytablez.Fields("serie") = "" & mytabler.Fields("serie")
            mytablez.Fields("acu") = "B"

        End If

        If "" & mytabler.Fields("tipo") = "4" Then  'factura
            mytablez.Fields("serie") = "" & mytabler.Fields("serie")
            mytablez.Fields("numero") = "" & mytabler.Fields("numero")
            mytablez.Fields("acu") = "A"

        End If

        If "" & mytabler.Fields("tipo") = "5" Then  'factura
            mytablez.Fields("serie") = "N" & mytabler.Fields("caja")
            mytablez.Fields("numero") = Mid$("" & mytabler.Fields("numero"), 5, 11)
            mytablez.Fields("acu") = "G"

        End If

    End If
   
    If "" & mytabler.Fields("acu") = "X" Or "" & mytabler.Fields("acu") = "Y" Then
        mytablez.Fields("tipoclie") = "C"

        If "" & mytabler.Fields("acu") = "X" Then
            mytablez.Fields("tipoclie") = "W"

        End If

        If "" & mytabler.Fields("acu") = "Y" Then
            mytablez.Fields("tipoclie") = "V"

        End If

        GoTo ay1

    End If
   
    If "" & mytabler.Fields("acu") = "V" Then
        mytablez.Fields("tipoclie") = "C"

        If "" & mytabler.Fields("tipo") = "1" Then  'ticket boleta
            mytablez.Fields("acu") = "C"

        End If

        If "" & mytabler.Fields("tipo") = "2" Then  'ticket factura
            mytablez.Fields("acu") = "D"

        End If

        If "" & mytabler.Fields("tipo") = "3" Then  'boleta
            mytablez.Fields("acu") = "B"

        End If

        If "" & mytabler.Fields("tipo") = "4" Then  'factura
            mytablez.Fields("acu") = "A"

        End If

        If "" & mytabler.Fields("tipo") = "5" Then  'factura
            mytablez.Fields("acu") = "G"

        End If

    End If
   
    If "" & mytabler.Fields("acu") = "C" Then
        mytablez.Fields("tipoclie") = "P"

        If "" & mytabler.Fields("tipo") = "CI" Then  'INTERNO
            mytablez.Fields("acu") = "P"

        End If

        If "" & mytabler.Fields("tipo") = "CA" Then  'FACTURA COMPRA
            mytablez.Fields("acu") = "K"

        End If

        If "" & mytabler.Fields("tipo") = "CB" Then  'BOLETA COMPRA
            mytablez.Fields("acu") = "J"

        End If
   
    End If
   
ay1:
   
    mytablez.Fields("codigo") = "" & mytabler.Fields("codigo")
    mytablez.Fields("paridad") = Val("" & mytabler.Fields("paridadfp"))
    mytablez.Fields("codigo") = "" & mytabler.Fields("codigo")
    mytablez.Fields("nombre") = "" & mytabler.Fields("nombrefp")
    'mytablez.Fields("tipo") = "" & mytabler.Fields("tipo")
    'mytablez.Fields("serie") = "" & mytabler.Fields("caja")
    'mytablez.Fields("numero") = Mid$("" & mytabler.Fields("numero"), 1, 11)
    mytablez.Fields("tipoclie") = "C"
    'mytabley.Fields("codigo") = "" & xruc
    mytablez.Fields("fecha") = "" & mytabler.Fields("fecha")
    mytablez.Fields("moneda") = "" & mytabler.Fields("monedafp")
    mytablez.Fields("total") = Val("" & mytabler.Fields("total"))
   
    mytablez.Fields("caja") = "" & mytabler.Fields("caja")
    mytablez.Fields("turno") = "" & mytabler.Fields("turno")
    mytablez.Fields("usuario") = "" & mytabler.Fields("usuario")
   
    mytablez.Fields("total") = Val("" & mytabler.Fields("total"))
    mytablez.Fields("cambio") = Val("" & mytabler.Fields("cambio"))
    mytablez.Fields("recibe") = Val("" & mytabler.Fields("valorpagad"))
    mytablez.Fields("recibes") = Val("" & mytabler.Fields("vueltos"))
    mytablez.Fields("recibed") = Val("" & mytabler.Fields("vueltod"))
    mytablez.Fields("saldos") = 0 'Val("" & mytabler.Fields("saldos"))
    mytablez.Fields("saldod") = 0 'Val("" & mytabler.Fields("saldod"))
    mytablez.Fields("nombre") = "" & mytabler.Fields("nombre")
    'mytablez.Fields("orden") = "" & mytabler.Fields("orden")
    mytablez.Fields("observa") = Mid$("" & mytabler.Fields("observacio"), 1, 15)
    mytablez.Fields("dias") = "" & mytabler.Fields("dias")
    mytablez.Fields("fpago") = "" & mytabler.Fields("fpago")
    'mytablez.Fields("acufp") = "" '& mytabler.Fields("acufp")
    mytablez.Fields("descripcio") = "" & mytabler.Fields("observacio")
    'mytablez.Fields("acu") = "" & mytabler.Fields("acu")
    mytablez.Fields("local") = "01" '& mytabler.Fields("local")
    mytablez.Fields("servicio") = "" & mytabler.Fields("servicio")

End Sub

Sub pone_registro_ventax(mytablez As ADODB.Recordset, mytabler As Snapshot)

    mytablez.Fields("estado") = "" & mytabler.Fields("estado")
    mytablez.Fields("tipo") = "" & mytabler.Fields("tipo")
    mytablez.Fields("serie") = "" & mytabler.Fields("serie")
    mytablez.Fields("numero") = Mid$("" & mytabler.Fields("numero"), 1, 11)
   
    If "" & mytabler.Fields("acu") = "V" Then
        mytablez.Fields("tipoclie") = "C"

        If "" & mytabler.Fields("tipo") = "1" Then  'ticket boleta
            mytablez.Fields("serie") = "B" & mytabler.Fields("caja")
            mytablez.Fields("numero") = Mid$("" & mytabler.Fields("numero"), 5, 11)
            mytablez.Fields("acu") = "C"

        End If

        If "" & mytabler.Fields("tipo") = "2" Then  'ticket factura
            mytablez.Fields("serie") = "F" & mytabler.Fields("caja")
            mytablez.Fields("numero") = Mid$("" & mytabler.Fields("numero"), 5, 11)
            mytablez.Fields("acu") = "D"

        End If

        If "" & mytabler.Fields("tipo") = "3" Then  'boleta
            mytablez.Fields("numero") = Mid$("" & mytabler.Fields("numero"), 5, 11)
            mytablez.Fields("serie") = "" & mytabler.Fields("serie")
            mytablez.Fields("acu") = "B"

        End If

        If "" & mytabler.Fields("tipo") = "4" Then  'factura
            mytablez.Fields("serie") = "" & mytabler.Fields("serie")
            mytablez.Fields("numero") = "" & mytabler.Fields("numero")
            mytablez.Fields("acu") = "A"

        End If

        If "" & mytabler.Fields("tipo") = "5" Then  'factura
            mytablez.Fields("serie") = "N" & mytabler.Fields("caja")
            mytablez.Fields("numero") = Mid$("" & mytabler.Fields("numero"), 5, 11)
            mytablez.Fields("acu") = "G"

        End If

    End If
   
    'If "" & mytabler.Fields("acu") = "V" Then
    'mytablez.Fields("tipoclie") = "C"
    'If "" & mytabler.Fields("tipo") = "1" Then  'ticket boleta
    '   mytablez.Fields("acu") = "C"
    'End If
    'If "" & mytabler.Fields("tipo") = "2" Then  'ticket factura
    '   mytablez.Fields("acu") = "D"
    'End If
    'If "" & mytabler.Fields("tipo") = "3" Then  'boleta
    '   mytablez.Fields("acu") = "B"
    'End If
    'If "" & mytabler.Fields("tipo") = "4" Then  'factura
    '   mytablez.Fields("acu") = "A"
    'End If
    'If "" & mytabler.Fields("tipo") = "5" Then  'factura
    '   mytablez.Fields("acu") = "G"
    'End If
    'End If
   
    'If "" & mytabler.Fields("acu") = "C" Then
    'mytablez.Fields("tipoclie") = "P"
    'If "" & mytabler.Fields("tipo") = "CI" Then  'INTERNO
    '   mytablez.Fields("acu") = "P"
    'End If
    'If "" & mytabler.Fields("tipo") = "CA" Then  'FACTURA COMPRA
    '   mytablez.Fields("acu") = "K"
    'End If
    'If "" & mytabler.Fields("tipo") = "CB" Then  'BOLETA COMPRA
    '   mytablez.Fields("acu") = "J"
    'End If
    '
    ' End If
   
    mytablez.Fields("codigo") = "" & mytabler.Fields("codigo")
    mytablez.Fields("acu1") = "" & mytabler.Fields("acu1")
    mytablez.Fields("fecha") = "" & mytabler.Fields("fecha")
    mytablez.Fields("moneda") = "" & mytabler.Fields("moneda")
    mytablez.Fields("producto") = "" & mytabler.Fields("producto")
    mytablez.Fields("descripcio") = "" & mytabler.Fields("descripcio")
    mytablez.Fields("unidad") = "" & mytabler.Fields("unidad")
    mytablez.Fields("factor") = Val("" & mytabler.Fields("factor"))
    mytablez.Fields("cantidad") = Val("" & mytabler.Fields("cantidad"))
    mytablez.Fields("precio") = Val("" & mytabler.Fields("precio"))
    mytablez.Fields("igv") = Val("" & mytabler.Fields("igv"))
    mytablez.Fields("neto") = Val("" & mytabler.Fields("bruto"))
    mytablez.Fields("descuento") = Val("" & mytabler.Fields("descuento"))
    mytablez.Fields("subtotal") = Val("" & mytabler.Fields("subtotal"))
    mytablez.Fields("impuesto") = Val("" & mytabler.Fields("impuesto"))
    mytablez.Fields("total") = Val("" & mytabler.Fields("total"))
    mytablez.Fields("estado") = "" & mytabler.Fields("estado")
    mytablez.Fields("fechacrea") = "" & mytabler.Fields("fecha")
    mytablez.Fields("hora") = "" & mytabler.Fields("hora")
    mytablez.Fields("vendedor") = "" & mytabler.Fields("vendedor")
    mytablez.Fields("bodega") = almacen
    mytablez.Fields("bodega") = "" & mytabler.Fields("bodega")
    mytablez.Fields("bodegaf") = ""
    mytablez.Fields("deslipo") = Val("" & mytabler.Fields("deslipo"))
    mytablez.Fields("flage") = ""
    mytablez.Fields("linea") = ""
    mytablez.Fields("t1") = 0
    mytablez.Fields("t2") = 0
    mytablez.Fields("t3") = 0
    mytablez.Fields("t4") = 0
    mytablez.Fields("t5") = 0
    mytablez.Fields("t6") = 0
    mytablez.Fields("t7") = 0
    mytablez.Fields("t8") = 0
    mytablez.Fields("t9") = 0
    mytablez.Fields("t10") = 0
    mytablez.Fields("t11") = 0
    mytablez.Fields("t12") = 0
    mytablez.Fields("t13") = 0
    mytablez.Fields("t14") = 0
    mytablez.Fields("t15") = 0
    mytablez.Fields("t16") = 0
    mytablez.Fields("l1") = ""
    mytablez.Fields("l2") = ""
    mytablez.Fields("l3") = ""
    mytablez.Fields("l4") = ""
    mytablez.Fields("local") = ""
    mytablez.Fields("proveedorp") = ""
    mytablez.Fields("observa1") = ""
    mytablez.Fields("observa2") = ""
    mytablez.Fields("observa3") = ""
    mytablez.Fields("observa4") = ""
    mytablez.Fields("zona") = ""
    mytablez.Fields("isc") = Val("" & mytabler.Fields("isc"))
    mytablez.Fields("tax") = 0
    mytablez.Fields("vtaneta") = 0
    mytablez.Fields("tcosto") = 0
    mytablez.Fields("ganancia") = 0
    mytablez.Fields("comision") = 0
    mytablez.Fields("usuario") = "" & mytabler.Fields("usuario")
    mytablez.Fields("caja") = "" & mytabler.Fields("caja")
    mytablez.Fields("turno") = "" & mytabler.Fields("turno")
    mytablez.Fields("servicio") = "" & mytabler.Fields("servicio")
    mytablez.Fields("comanda") = ""
    mytablez.Fields("mesa") = ""
    mytablez.Fields("salon") = ""
    mytablez.Fields("mesero") = ""
    'mytablez.Fields("codigop") = ""
    mytablez.Fields("local") = "01" '& mytabler.Fields("local")

End Sub

Sub pone_registro_venta1(mytablex As Table, mytabley As Snapshot, buf1 As String)

    Dim buf As String

    mytablex.Fields("acu") = "V"
    mytablex.Fields("tipo") = "" & mytabley.Fields("tipo")
    mytablex.Fields("numero") = buf1 '"" & mytabley.Fields("numero")
    mytablex.Fields("serie") = Mid$("" & mytabley.Fields("caja"), 1, 3)
    mytablex.Fields("codigo") = "" & mytabley.Fields("codigo")
    mytablex.Fields("fecha") = mytabley.Fields("fecha")
    mytablex.Fields("fechav") = mytabley.Fields("fecha")
    mytablex.Fields("moneda") = "" & mytabley.Fields("moneda")
    mytablex.Fields("vendedor") = "" & mytabley.Fields("vendedor")
    mytablex.Fields("fpago") = "" & mytabley.Fields("fpago")
    mytablex.Fields("paridad") = Val("" & mytabley.Fields("paridad"))
    mytablex.Fields("dias") = Val("" & mytabley.Fields("dias"))
    mytablex.Fields("bodega") = Val("" & mytabley.Fields("bodega"))
    mytablex.Fields("bodegaf") = ""
    mytablex.Fields("estado") = "" & mytabley.Fields("estado")
    mytablex.Fields("acu1") = ""
    mytablex.Fields("usuario") = "" & mytabley.Fields("usuario")
    mytablex.Fields("hora") = "" & mytabley.Fields("hora")
    mytablex.Fields("nombreb") = "" & mytabley.Fields("nombre")
    mytablex.Fields("total") = Val("" & mytabley.Fields("total"))
    mytablex.Fields("descuento") = Val("" & mytabley.Fields("descuento"))
    mytablex.Fields("impuesto") = Val("" & mytabley.Fields("impuesto"))
    mytablex.Fields("subtotal") = Val("" & mytabley.Fields("subtotal"))
    mytablex.Fields("local") = "02" '& mytabley.Fields("local")
    mytablex.Fields("usuario") = "" & mytabley.Fields("usuario")
    mytablex.Fields("caja") = "" & mytabley.Fields("caja")
    mytablex.Fields("turno") = "" & mytabley.Fields("turno")
    mytablex.Fields("servicio") = "" & mytabley.Fields("servicio")
    mytablex.Fields("comanda") = "" & mytabley.Fields("comanda")
    mytablex.Fields("mesa") = "" & mytabley.Fields("mesa")
    mytablex.Fields("salon") = "" & mytabley.Fields("salon")
    mytablex.Fields("vendedor") = "" & mytabley.Fields("vendedor")
    mytablex.Fields("ruc") = "" & mytabley.Fields("codigo")
    'mytablex.Fields("estado") = "" & mytabley.Fields("estado")

End Sub

Sub pone_registro_venta(mytablex As ADODB.Recordset, mytabley As Snapshot)
   
    If "" & mytabley.Fields("acu") = "V" Then
        mytablex.Fields("tipoclie") = "C"

        If "" & mytabley.Fields("tipo") = "1" Then  'ticket boleta
            mytablex.Fields("serie") = "B" & mytabley.Fields("caja")
            mytablex.Fields("numero") = Mid$("" & mytabley.Fields("numero"), 5, 11)
            mytablex.Fields("acu") = "C"

        End If

        If "" & mytabley.Fields("tipo") = "2" Then  'ticket factura
            mytablex.Fields("serie") = "F" & mytabley.Fields("caja")
            mytablex.Fields("numero") = Mid$("" & mytabley.Fields("numero"), 5, 11)
            mytablex.Fields("acu") = "D"

        End If

        If "" & mytabley.Fields("tipo") = "3" Then  'boleta
            mytablex.Fields("numero") = Mid$("" & mytabley.Fields("numero"), 5, 11)
            mytablex.Fields("serie") = "" & mytabley.Fields("serie")
            mytablex.Fields("acu") = "B"

        End If

        If "" & mytabley.Fields("tipo") = "4" Then  'factura
            mytablex.Fields("serie") = "" & mytabley.Fields("serie")
            mytablex.Fields("numero") = "" & mytabley.Fields("numero")
            mytablex.Fields("acu") = "A"

        End If

        If "" & mytabley.Fields("tipo") = "5" Then  'factura
            mytablex.Fields("serie") = "N" & mytabley.Fields("caja")
            mytablex.Fields("numero") = Mid$("" & mytabley.Fields("numero"), 5, 11)
            mytablex.Fields("acu") = "G"

        End If

    End If
   
    'If "" & mytabley.Fields("acu") = "C" Then
    'mytablex.Fields("tipoclie") = "P"
    'If "" & mytabley.Fields("tipo") = "CI" Then  'INTERNO
    '   mytablex.Fields("acu") = "P"
    'End If
    'If "" & mytabley.Fields("tipo") = "CA" Then  'FACTURA COMPRA
    '   mytablex.Fields("acu") = "K"
    'End If
    'If "" & mytabley.Fields("tipo") = "CB" Then  'BOLETA COMPRA
    '   mytablex.Fields("acu") = "J"
    'End If
    '
    'End If

    mytablex.Fields("tipo") = "" & mytabley.Fields("tipo")
    'mytablex.Fields("serie") = "" & mytabley.Fields("serie")

    mytablex.Fields("codigo") = "" & mytabley.Fields("codigo")
    mytablex.Fields("partida") = ""
    mytablex.Fields("destino") = ""
    mytablex.Fields("fecha") = mytabley.Fields("fecha")
    mytablex.Fields("fechae") = mytabley.Fields("fecha")
    mytablex.Fields("moneda") = "" & mytabley.Fields("moneda")
    mytablex.Fields("vendedor") = "" & mytabley.Fields("vendedor")
    'mytablex.Fields("transporte") = "" & mytabley.Fields("transporte")
    mytablex.Fields("fpago") = "" & mytabley.Fields("fpago")
    mytablex.Fields("paridad") = Val("" & mytabley.Fields("paridad"))
    mytablex.Fields("dias") = Val("" & mytabley.Fields("dias"))
    mytablex.Fields("bodega") = almacen
    mytablex.Fields("bodega") = "" & mytabley.Fields("bodega")
    mytablex.Fields("bodegaf") = ""
    'mytablex.Fields("observa") = "" & mytabley.Fields("observa")
    mytablex.Fields("estado") = "" & mytabley.Fields("estado")
    mytablex.Fields("acu1") = ""
    mytablex.Fields("usuario") = "" & mytabley.Fields("usuario")
    mytablex.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
    mytablex.Fields("hora") = "" & mytabley.Fields("hora")
    mytablex.Fields("nombre") = "" & mytabley.Fields("nombreb")
    mytablex.Fields("total") = Val("" & mytabley.Fields("total"))
    mytablex.Fields("descuento") = Val("" & mytabley.Fields("descuento"))
    'mytablex.Fields("neto") = Val("" & mytabley.Fields("neto"))
    mytablex.Fields("impuesto") = Val("" & mytabley.Fields("impuesto"))
    mytablex.Fields("subtotal") = Val("" & mytabley.Fields("subtotal"))
    mytablex.Fields("flage") = ""
    mytablex.Fields("tipo1") = ""
    mytablex.Fields("serie1") = ""
    mytablex.Fields("numero1") = ""
    mytablex.Fields("serie2") = ""
    mytablex.Fields("numero2") = ""
    mytablex.Fields("serie3") = ""
    mytablex.Fields("numero3") = ""
    mytablex.Fields("serie4") = ""
    mytablex.Fields("numero4") = ""
    mytablex.Fields("serie5") = ""
    mytablex.Fields("numero5") = ""
    mytablex.Fields("serie6") = ""
    mytablex.Fields("numero6") = ""
    mytablex.Fields("serie7") = ""
    mytablex.Fields("numero7") = ""
    mytablex.Fields("serie8") = ""
    mytablex.Fields("numero8") = ""
    mytablex.Fields("nop") = ""
    mytablex.Fields("local") = "02"  '& mytabley.Fields("local")
    'mytablex.Fields("c1") = ""
    'mytablex.Fields("c1") = ""
    'mytablex.Fields("c1") = ""
    'mytablex.Fields("c1") = ""
    mytablex.Fields("zona") = ""
    mytablex.Fields("retipo1") = ""
    mytablex.Fields("renumero1") = ""
    mytablex.Fields("renumero2") = ""
    mytablex.Fields("renumero3") = ""
    mytablex.Fields("retotal") = 0
    mytablex.Fields("retotal1") = 0
    'mytablex.Fields("retota2") = 0
    mytablex.Fields("retotal3") = 0
    mytablex.Fields("retotal") = 0
    'mytablex.Fields("acuenta") = ""
    mytablex.Fields("retotal") = 0
    mytablex.Fields("nro_items") = 0
    mytablex.Fields("retotal") = 0
    mytablex.Fields("adetotal") = 0
    mytablex.Fields("retotal") = 0
    mytablex.Fields("yausado") = ""
    mytablex.Fields("retotal") = 0
    mytablex.Fields("usuario") = "" & mytabley.Fields("usuario")
    mytablex.Fields("caja") = "" & mytabley.Fields("caja")
    mytablex.Fields("turno") = "" & mytabley.Fields("turno")
    mytablex.Fields("servicio") = "A" '& mytabley.Fields("servicio")
    mytablex.Fields("comanda") = "" & mytabley.Fields("comanda")
    mytablex.Fields("mesa") = "" & mytabley.Fields("mesa")
    mytablex.Fields("salon") = "" & mytabley.Fields("salon")
    mytablex.Fields("mesero") = "" & mytabley.Fields("vendedor")
    'mytablex.Fields("telefono") = "" & mytabley.Fields("telefono")
    mytablex.Fields("ruc") = "" & mytabley.Fields("ruc")
    mytablex.Fields("montopagar") = 0
    mytablex.Fields("tdocdeli") = ""
    mytablex.Fields("gravado") = 0
    mytablex.Fields("fechasunat") = mytabley.Fields("fecha")

End Sub

Sub pone_registro_pedido(mytablex As Table, mytabley As Snapshot)
    mytablex.Fields("tipoclie") = "C"
    mytablex.Fields("acu") = "I"
    mytablex.Fields("tipo") = "" & mytabley.Fields("tipo")
    mytablex.Fields("serie") = "" & mytabley.Fields("serie")
    mytablex.Fields("numero") = Mid$("" & mytabley.Fields("numero"), 1, 11)
    mytablex.Fields("codigo") = "" & mytabley.Fields("codigo")
    mytablex.Fields("partida") = ""
    mytablex.Fields("destino") = ""
    mytablex.Fields("fecha") = mytabley.Fields("fecha")
    mytablex.Fields("fechae") = mytabley.Fields("fecha")
    mytablex.Fields("moneda") = "" & mytabley.Fields("moneda")
    mytablex.Fields("vendedor") = "" & mytabley.Fields("vendedor")
    'mytablex.Fields("transporte") = "" & mytabley.Fields("transporte")
    mytablex.Fields("fpago") = "" & mytabley.Fields("fpago")
    mytablex.Fields("paridad") = Val("" & mytabley.Fields("paridad"))
    mytablex.Fields("dias") = Val("" & mytabley.Fields("dias"))
    mytablex.Fields("bodega") = almacen
    mytablex.Fields("bodega") = "" & mytabley.Fields("bodega")
    mytablex.Fields("bodegaf") = ""
    'mytablex.Fields("observa") = "" & mytabley.Fields("observa")
    mytablex.Fields("estado") = "" & mytabley.Fields("estado")
    mytablex.Fields("acu1") = ""
    mytablex.Fields("usuario") = "" & mytabley.Fields("usuario")
    mytablex.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
    mytablex.Fields("hora") = "" & mytabley.Fields("hora")
    mytablex.Fields("nombre") = "" & mytabley.Fields("nombreb")
    mytablex.Fields("total") = Val("" & mytabley.Fields("total"))
    mytablex.Fields("descuento") = Val("" & mytabley.Fields("descuento"))
    'mytablex.Fields("neto") = Val("" & mytabley.Fields("neto"))
    mytablex.Fields("impuesto") = Val("" & mytabley.Fields("impuesto"))
    mytablex.Fields("subtotal") = Val("" & mytabley.Fields("subtotal"))
    mytablex.Fields("flage") = ""
    mytablex.Fields("tipo1") = ""
    mytablex.Fields("serie1") = ""
    mytablex.Fields("numero1") = ""
    mytablex.Fields("serie2") = ""
    mytablex.Fields("numero2") = ""
    mytablex.Fields("serie3") = ""
    mytablex.Fields("numero3") = ""
    mytablex.Fields("serie4") = ""
    mytablex.Fields("numero4") = ""
    mytablex.Fields("serie5") = ""
    mytablex.Fields("numero5") = ""
    mytablex.Fields("serie6") = ""
    mytablex.Fields("numero6") = ""
    mytablex.Fields("serie7") = ""
    mytablex.Fields("numero7") = ""
    mytablex.Fields("serie8") = ""
    mytablex.Fields("numero8") = ""
    mytablex.Fields("nop") = ""
    mytablex.Fields("local") = "02" '& mytabley.Fields("local")
    'mytablex.Fields("c1") = ""
    'mytablex.Fields("c1") = ""
    'mytablex.Fields("c1") = ""
    'mytablex.Fields("c1") = ""
    mytablex.Fields("zona") = ""
    mytablex.Fields("retipo1") = ""
    mytablex.Fields("renumero1") = ""
    mytablex.Fields("renumero2") = ""
    mytablex.Fields("renumero3") = ""
    mytablex.Fields("retotal") = 0
    mytablex.Fields("retotal1") = 0
    'mytablex.Fields("retota2") = 0
    mytablex.Fields("retotal3") = 0
    mytablex.Fields("retotal") = 0
    'mytablex.Fields("acuenta") = ""
    mytablex.Fields("retotal") = 0
    mytablex.Fields("nro_items") = 0
    mytablex.Fields("retotal") = 0
    mytablex.Fields("adetotal") = 0
    mytablex.Fields("retotal") = 0
    mytablex.Fields("yausado") = ""
    mytablex.Fields("retotal") = 0
    mytablex.Fields("usuario") = "" & mytabley.Fields("usuario")
    mytablex.Fields("caja") = "" & mytabley.Fields("caja")
    mytablex.Fields("turno") = "" & mytabley.Fields("turno")
    mytablex.Fields("servicio") = "" & mytabley.Fields("servicio")
    mytablex.Fields("comanda") = "" & mytabley.Fields("comanda")
    mytablex.Fields("mesa") = "" & mytabley.Fields("mesa")
    mytablex.Fields("salon") = "" & mytabley.Fields("salon")
    mytablex.Fields("mesero") = "" & mytabley.Fields("vendedor")
    'mytablex.Fields("telefono") = "" & mytabley.Fields("telefono")
    mytablex.Fields("ruc") = "" & mytabley.Fields("ruc")
    mytablex.Fields("montopagar") = 0
    mytablex.Fields("tdocdeli") = ""
    mytablex.Fields("gravado") = 0
    mytablex.Fields("fechasunat") = mytabley.Fields("fecha")

End Sub

Sub pone_registro_ocompra(mytablex As Table, mytabley As Snapshot)
    mytablex.Fields("tipoclie") = "P"
    mytablex.Fields("acu") = "R"
    mytablex.Fields("tipo") = "" & mytabley.Fields("tipo")
    mytablex.Fields("serie") = "" & mytabley.Fields("serie")
    mytablex.Fields("numero") = Mid$("" & mytabley.Fields("numero"), 1, 11)
    mytablex.Fields("codigo") = "" & mytabley.Fields("codigo")
    mytablex.Fields("partida") = ""
    mytablex.Fields("destino") = ""
    mytablex.Fields("fecha") = mytabley.Fields("fecha")
    mytablex.Fields("fechae") = mytabley.Fields("fecha")
    mytablex.Fields("moneda") = "" & mytabley.Fields("moneda")
    mytablex.Fields("vendedor") = "" & mytabley.Fields("vendedor")
    'mytablex.Fields("transporte") = "" & mytabley.Fields("transporte")
    mytablex.Fields("fpago") = "" & mytabley.Fields("fpago")
    mytablex.Fields("paridad") = Val("" & mytabley.Fields("paridad"))
    mytablex.Fields("dias") = Val("" & mytabley.Fields("dias"))
    mytablex.Fields("bodega") = almacen
    mytablex.Fields("bodega") = "" & mytabley.Fields("bodega")
    mytablex.Fields("bodegaf") = ""
    'mytablex.Fields("observa") = "" & mytabley.Fields("observa")
    mytablex.Fields("estado") = "" & mytabley.Fields("estado")
    mytablex.Fields("acu1") = ""
    mytablex.Fields("usuario") = "" & mytabley.Fields("usuario")
    mytablex.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
    mytablex.Fields("hora") = "" & mytabley.Fields("hora")
    mytablex.Fields("nombre") = "" & mytabley.Fields("nombreb")
    mytablex.Fields("total") = Val("" & mytabley.Fields("total"))
    mytablex.Fields("descuento") = Val("" & mytabley.Fields("descuento"))
    'mytablex.Fields("neto") = Val("" & mytabley.Fields("neto"))
    mytablex.Fields("impuesto") = Val("" & mytabley.Fields("impuesto"))
    mytablex.Fields("subtotal") = Val("" & mytabley.Fields("subtotal"))
    mytablex.Fields("flage") = ""
    mytablex.Fields("tipo1") = ""
    mytablex.Fields("serie1") = ""
    mytablex.Fields("numero1") = ""
    mytablex.Fields("serie2") = ""
    mytablex.Fields("numero2") = ""
    mytablex.Fields("serie3") = ""
    mytablex.Fields("numero3") = ""
    mytablex.Fields("serie4") = ""
    mytablex.Fields("numero4") = ""
    mytablex.Fields("serie5") = ""
    mytablex.Fields("numero5") = ""
    mytablex.Fields("serie6") = ""
    mytablex.Fields("numero6") = ""
    mytablex.Fields("serie7") = ""
    mytablex.Fields("numero7") = ""
    mytablex.Fields("serie8") = ""
    mytablex.Fields("numero8") = ""
    mytablex.Fields("nop") = ""
    mytablex.Fields("local") = "02" '& mytabley.Fields("local")
    'mytablex.Fields("c1") = ""
    'mytablex.Fields("c1") = ""
    'mytablex.Fields("c1") = ""
    'mytablex.Fields("c1") = ""
    mytablex.Fields("zona") = ""
    mytablex.Fields("retipo1") = ""
    mytablex.Fields("renumero1") = ""
    mytablex.Fields("renumero2") = ""
    mytablex.Fields("renumero3") = ""
    mytablex.Fields("retotal") = 0
    mytablex.Fields("retotal1") = 0
    'mytablex.Fields("retota2") = 0
    mytablex.Fields("retotal3") = 0
    mytablex.Fields("retotal") = 0
    'mytablex.Fields("acuenta") = ""
    mytablex.Fields("retotal") = 0
    mytablex.Fields("nro_items") = 0
    mytablex.Fields("retotal") = 0
    mytablex.Fields("adetotal") = 0
    mytablex.Fields("retotal") = 0
    mytablex.Fields("yausado") = ""
    mytablex.Fields("retotal") = 0
    mytablex.Fields("usuario") = "" & mytabley.Fields("usuario")
    mytablex.Fields("caja") = "" & mytabley.Fields("caja")
    mytablex.Fields("turno") = "" & mytabley.Fields("turno")
    mytablex.Fields("servicio") = "" & mytabley.Fields("servicio")
    mytablex.Fields("comanda") = "" & mytabley.Fields("comanda")
    mytablex.Fields("mesa") = "" & mytabley.Fields("mesa")
    mytablex.Fields("salon") = "" & mytabley.Fields("salon")
    mytablex.Fields("mesero") = "" & mytabley.Fields("vendedor")
    'mytablex.Fields("telefono") = "" & mytabley.Fields("telefono")
    mytablex.Fields("ruc") = "" & mytabley.Fields("ruc")
    mytablex.Fields("montopagar") = 0
    mytablex.Fields("tdocdeli") = ""
    mytablex.Fields("gravado") = 0
    mytablex.Fields("fechasunat") = mytabley.Fields("fecha")

End Sub

Private Sub procesar_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    Command1_Click

End Sub

Sub proceso_importar()

    Dim mydby    As Database

    Dim mytablex As Table

    Dim mytablez As Table

    Dim mytablea As Table

    Dim mytableb As Table

    Dim mytablec As Table

    Dim mytabled As Table

    Dim mytabley As Snapshot

    Dim buf      As String

    Dim vr

    Dim ind As Long

    ind = 1
    xfecha = ""
    buf = "select * from accesori "
    Set mydby = OpenDatabase("\rp_orion.v2\maximo\", False, False, "foxpro 2.5;")
    Set mytabley = mydbxglo.CreateSnapshot(buf)

    Set mytablex = mydbxglo.OpenTable("producto")
    mytablex.Index = "producto"
    Set mytablez = mydbxglo.OpenTable("familia")
    mytablez.Index = "familia"
    Set mytablea = mydbxglo.OpenTable("subfamilia")
    mytablea.Index = "subfamilia"
    Set mytableb = mydbxglo.OpenTable("marca")
    mytableb.Index = "marca"
    Set mytablec = mydbxglo.OpenTable("almacen")
    mytablec.Index = "almacen"
    Set mytabled = mydbxglo.OpenTable("saldoini")
    mytabled.Index = "saldoini"

    Do

        If mytabley.EOF Then Exit Do
        'ind = ind + 1
        mytablex.Seek "=", "" & mytabley.Fields("producto")

        If mytablex.NoMatch Then
            mytablex.AddNew
            pone_registro_importa mytablex, mytabley, ind
            mytablex.Update

        End If

        If Not mytablex.NoMatch Then
            mytablex.Edit
            pone_registro_importa mytablex, mytabley, ind
            mytablex.Update

        End If

        'graba familias
        mytablez.Seek "=", Mid$("" & mytabley.Fields("familia"), 1, 6)

        If mytablez.NoMatch Then
            mytablez.AddNew
            mytablez.Fields("familia") = Mid$("" & mytabley.Fields("familia"), 1, 6)
            mytablez.Fields("descripcio") = Mid$("" & mytabley.Fields("familia"), 1, 15)
            mytablez.Update

        End If

        If Not mytablez.NoMatch Then
            mytablez.Edit
            mytablez.Fields("familia") = Mid$("" & mytabley.Fields("familia"), 1, 6)
            mytablez.Fields("descripcio") = Mid$("" & mytabley.Fields("familia"), 1, 15)
            mytablez.Update

        End If

        'graba almacenes
        graba_almacenes mytablec, mytabled, mytabley

        'graba subfamilias
        mytablea.Seek "=", Mid$("" & mytabley.Fields("subfamilia"), 1, 6)

        If mytablea.NoMatch Then
            mytablea.AddNew
            mytablea.Fields("subfamilia") = Mid$("" & mytabley.Fields("subfamilia"), 1, 6)
            mytablea.Fields("familia") = Mid$("" & mytabley.Fields("familia"), 1, 6)
            mytablea.Fields("descripcio") = Mid$("" & mytabley.Fields("subfamilia"), 1, 15)
            mytablea.Update

        End If

        If Not mytablea.NoMatch Then
            mytablea.Edit
            mytablea.Fields("subfamilia") = Mid$("" & mytabley.Fields("subfamilia"), 1, 6)
            mytablea.Fields("familia") = Mid$("" & mytabley.Fields("familia"), 1, 6)
            mytablea.Fields("descripcio") = Mid$("" & mytabley.Fields("subfamilia"), 1, 15)
            mytablea.Update

        End If

        'graba marca
        mytableb.Seek "=", Mid$("" & mytabley.Fields("marca"), 1, 6)

        If mytableb.NoMatch Then
            mytableb.AddNew
            mytableb.Fields("marca") = Mid$("" & mytabley.Fields("marca"), 1, 6)
            mytableb.Fields("descripcio") = Mid$("" & mytabley.Fields("marca"), 1, 15)
            mytableb.Update

        End If

        If Not mytableb.NoMatch Then
            mytableb.Edit
            mytableb.Fields("marca") = Mid$("" & mytabley.Fields("marca"), 1, 6)
            mytableb.Fields("descripcio") = Mid$("" & mytabley.Fields("marca"), 1, 15)
            mytableb.Update

        End If

        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    mytabley.Close
    mytablex.Close
 
    MsgBox "proceso Terminado", 48, "Aviso"

End Sub

Sub graba_almacenes(mytablea As Table, mytableb As Table, mytabley As Table)

    Dim I As Integer

    For I = 1 To 7
        mytablea.Seek "=", "" & mytabley.Fields("producto"), Format(I, "00")

        If mytablea.NoMatch Then
            mytablea.AddNew
            mytablea.Fields("producto") = "" & mytabley.Fields("producto")
            mytablea.Fields("bodega") = Format(I, "00")
            mytablea.Fields("saldo") = 0

            If I = 6 Then
                mytablea.Fields("saldo") = Val("" & mytabley.Fields("stock"))

            End If

            mytablea.Update

        End If

        If Not mytablea.NoMatch Then
            mytablea.Edit
            mytablea.Fields("producto") = "" & mytabley.Fields("producto")
            mytablea.Fields("bodega") = Format(I, "00")
            mytablea.Fields("saldo") = 0

            If I = 6 Then
                mytablea.Fields("saldo") = Val("" & mytabley.Fields("stock"))

            End If

            mytablea.Update

        End If

    Next I

    'saldoinicial
    mytableb.Seek "=", "" & mytabley.Fields("producto"), "06", Format(Now, "dd/mm/yyyy")

    If mytableb.NoMatch Then
        mytableb.AddNew
        mytableb.Fields("familia") = Mid$("" & mytabley.Fields("familia"), 1, 6)
        mytableb.Fields("producto") = "" & mytabley.Fields("producto")
        mytableb.Fields("DESCRIPCIO") = Trim("" & mytabley.Fields("familia")) + " " + Trim("" & mytabley.Fields("subfamilia")) + " " + Trim("" & mytabley.Fields("marca")) + " " + Mid$(Trim("" & mytabley.Fields("codfab")), 1, 35)
        mytableb.Fields("unidad") = "UND"
        mytableb.Fields("factor") = 1
        mytableb.Fields("bodega") = "06"
        mytableb.Fields("fecha") = Format(Now, "dd/mm/yyyy")
        mytableb.Fields("cantidad") = Val("" & mytabley.Fields("stock"))
        mytableb.Update

    End If

    If Not mytableb.NoMatch Then
        mytableb.Edit
        mytableb.Fields("familia") = Mid$("" & mytabley.Fields("familia"), 1, 6)
        mytableb.Fields("producto") = "" & mytabley.Fields("producto")
        mytableb.Fields("DESCRIPCIO") = Trim("" & mytabley.Fields("familia")) + " " + Trim("" & mytabley.Fields("subfamilia")) + " " + Trim("" & mytabley.Fields("marca")) + " " + Mid$(Trim("" & mytabley.Fields("codfab")), 1, 35)
        mytableb.Fields("unidad") = "UND"
        mytableb.Fields("factor") = 1
        mytableb.Fields("fecha") = Format(Now, "dd/mm/yyyy")
        mytableb.Fields("bodega") = "06"
        mytableb.Fields("cantidad") = Val("" & mytabley.Fields("stock"))
        mytableb.Update

    End If

End Sub

Sub pone_registro_importa(mytablex As Table, mytabley As Table, ind As Long)
    mytablex.Fields("producto") = "" & mytabley.Fields("producto")
    'mytablex.Fields("barras") = "" & mytabley.Fields("producto")
    mytablex.Fields("descripcio") = Trim("" & mytabley.Fields("familia")) + " " + Trim("" & mytabley.Fields("subfamilia")) + " " + Trim("" & mytabley.Fields("marca")) + " " + Mid$(Trim("" & mytabley.Fields("codfab")), 1, 35)
    mytablex.Fields("descorto") = Mid$(Trim("" & mytabley.Fields("familia")), 1, 6) + " " + Mid$(Trim("" & mytabley.Fields("subfamilia")), 1, 6) + " " + Mid$(Trim("" & mytabley.Fields("marca")), 1, 6)
    'mytablex.Fields("presenta") = "" & mytabley.Fields("presenta")
    mytablex.Fields("familia") = Mid$("" & mytabley.Fields("familia"), 1, 6)
    mytablex.Fields("subfamilia") = Mid$("" & mytabley.Fields("subfamilia"), 1, 6)
    mytablex.Fields("seccion") = ""
    mytablex.Fields("marca") = Mid$("" & mytabley.Fields("marca"), 1, 6)
    mytablex.Fields("categoria") = ""
    mytablex.Fields("linea") = ""
    mytablex.Fields("color") = ""
    mytablex.Fields("fabrica") = ""
    'mytablex.Fields("proveedor1") = "" & mytabley.Fields("proveedor")
    'mytablex.Fields("proveedor2") = ""
    'mytablex.Fields("codprov1") = "" & mytabley.Fields("cpa")
    'mytablex.Fields("codprov2") = ""
    mytablex.Fields("serie") = ""
    mytablex.Fields("peso") = ""
    mytablex.Fields("servicio") = ""
    mytablex.Fields("vtaund") = ""
    mytablex.Fields("oferta") = ""
    mytablex.Fields("vecaja") = "S"
    mytablex.Fields("igv") = 19
    mytablex.Fields("isc") = 0
    mytablex.Fields("pesokgr") = 0.001
    mytablex.Fields("comision") = 0
    mytablex.Fields("monedac") = "S"
    mytablex.Fields("unidad") = "UND"
    mytablex.Fields("factor") = "1"
    mytablex.Fields("costou") = 0
    mytablex.Fields("costop") = 0

    mytablex.Fields("monedav") = "S"
    mytablex.Fields("factor1") = "1"
    mytablex.Fields("unidad1") = "UND"
    mytablex.Fields("pventa1") = Val("" & mytabley.Fields("preuni1"))

    mytablex.Fields("factor2") = "1"
    mytablex.Fields("unidad2") = "UND"
    mytablex.Fields("pventa2") = Val("" & mytabley.Fields("preuni2"))

    mytablex.Fields("factor3") = "1"
    mytablex.Fields("unidad3") = "UND"
    mytablex.Fields("pventa3") = Val("" & mytabley.Fields("preuni3"))

    mytablex.Fields("factor4") = 12
    mytablex.Fields("unidad4") = "DOC"
    mytablex.Fields("pventa4") = Val("" & mytabley.Fields("predoc1"))

    mytablex.Fields("factor5") = 12
    mytablex.Fields("unidad5") = "DOC"
    mytablex.Fields("pventa5") = Val("" & mytabley.Fields("predoc2"))

End Sub

Sub temporal_graba()

    Dim mytablex As Table

    Dim mytabley As Table

    Dim mytablez As Table

    Dim mytablea As Table

    Set mytabley = mydbxglo.OpenTable("Producto")
    Set mytablex = mydbxglo.OpenTable("familia")
    Set mytablez = mydbxglo.OpenTable("subfamil")
    Set mytablea = mydbxglo.OpenTable("marca")
    mytablex.Index = "familia"
    mytablez.Index = "subfamilia"
    mytablea.Index = "marca"
    Do

        If mytabley.EOF Then Exit Do
        'familia
        mytablex.Seek "=", "" & mytabley.Fields("familia")

        If mytablex.NoMatch Then
            mytablex.AddNew
            mytablex.Fields("familia") = "" & mytabley.Fields("familia")
            mytablex.Fields("descripcio") = "" & mytabley.Fields("familia")
            mytablex.Update

        End If

        If Not mytablex.NoMatch Then
            mytablex.Edit
            mytablex.Fields("familia") = "" & mytabley.Fields("familia")
            mytablex.Fields("descripcio") = "" & mytabley.Fields("familia")
            mytablex.Update

        End If

        'subfamilia
        'mytablez.Seek "=", "" & mytabley.Fields("familia"), "" & mytabley.Fields("subfamilia")
        'If mytablez.NoMatch Then
        '   mytablez.AddNew
        '   mytablez.Fields("familia") = "" & mytabley.Fields("familia")
        '   mytablez.Fields("subfamilia") = "" & mytabley.Fields("subfamilia")
        '   mytablez.Fields("descripcio") = "" & mytabley.Fields("subfamilia")
        '   mytablez.Update
        'End If
        'If Not mytablez.NoMatch Then
        '   mytablez.Edit
        '   mytablez.Fields("familia") = "" & mytabley.Fields("familia")
        '   mytablez.Fields("subfamilia") = "" & mytabley.Fields("subfamilia")
        '   mytablez.Fields("descripcio") = "" & mytabley.Fields("subfamilia")
        '   mytablez.Update
        'End If
        'marca
        mytablea.Seek "=", "" & mytabley.Fields("marca")

        If mytablea.NoMatch Then
            mytablea.AddNew
            mytablea.Fields("marca") = "" & mytabley.Fields("marca")
            mytablea.Fields("descripcio") = "" & mytabley.Fields("marca")
            mytablea.Update

        End If

        If Not mytablea.NoMatch Then
            mytablea.Edit
            mytablea.Fields("marca") = "" & mytabley.Fields("marca")
            mytablea.Fields("descripcio") = "" & mytabley.Fields("marca")
            mytablea.Update

        End If

        mytabley.MoveNext
    Loop

    mytablex.Close
    MsgBox "proceso Terminado", 48, "Aviso"

End Sub

Sub importa_cajamarca()

    Dim mydby    As Database

    Dim mytablex As Table

    Dim mytablez As Table

    Dim mytablea As Table

    Dim mytableb As Table

    Dim mytablec As Table

    Dim mytabled As Table

    Dim mytabley As Snapshot

    Dim buf      As String

    Dim vr

    Dim ind As Long

    ind = 1
    xfecha = ""

    Set mydby = OpenDatabase("c:\URANIO\X", False, False, "foxpro 2.5;")
    Set mytabley = mydby.CreateSnapshot("OMG-ART")

    Set mytablex = mydbxglo.OpenTable("producto")
    mytablex.Index = "producto"
    Set mytablez = mydbxglo.OpenTable("familia")
    mytablez.Index = "familia"
    Set mytableb = mydbxglo.OpenTable("seccion")
    mytableb.Index = "seccion"
    Set mytablec = mydbxglo.OpenTable("codprov")
    mytablec.Index = "codprov"

    Do

        If mytabley.EOF Then Exit Do
        mytablex.Seek "=", "" & mytabley.Fields("plu")

        If mytablex.NoMatch Then
            mytablex.AddNew
            pone_registro_cajamarca mytablex, mytabley, mytablec
            mytablex.Update

        End If

        If Not mytablex.NoMatch Then
            mytablex.Edit
            pone_registro_cajamarca mytablex, mytabley, mytablec
            mytablex.Update

        End If

        'graba familias
        mytablez.Seek "=", Mid$("" & mytabley.Fields("familia"), 1, 6)

        If mytablez.NoMatch Then
            mytablez.AddNew
            mytablez.Fields("familia") = Mid$("" & mytabley.Fields("familia"), 1, 6)
            mytablez.Fields("descripcio") = Mid$("" & mytabley.Fields("familia"), 1, 15)
            mytablez.Update

        End If

        If Not mytablez.NoMatch Then
            mytablez.Edit
            mytablez.Fields("familia") = Mid$("" & mytabley.Fields("familia"), 1, 6)
            mytablez.Fields("descripcio") = Mid$("" & mytabley.Fields("familia"), 1, 15)
            mytablez.Update

        End If

        'graba seccion
        mytableb.Seek "=", Mid$("" & mytabley.Fields("ubicacio"), 1, 6)

        If mytableb.NoMatch Then
            mytableb.AddNew
            mytableb.Fields("seccion") = Mid$("" & mytabley.Fields("ubicacio"), 1, 6)
            mytableb.Fields("descripcio") = Mid$("" & mytabley.Fields("ubicacio"), 1, 15)
            mytableb.Update

        End If

        If Not mytableb.NoMatch Then
            mytableb.Edit
            mytableb.Fields("seccion") = Mid$("" & mytabley.Fields("ubicacio"), 1, 6)
            mytableb.Fields("descripcio") = Mid$("" & mytabley.Fields("ubicacio"), 1, 15)
            mytableb.Update

        End If

        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    mytabley.Close
    mytablex.Close
    MsgBox "proceso Terminado", 48, "Aviso"

End Sub

Sub pone_registro_cajamarca(mytablex As Table, mytabley As Table, mytablez As Table)
    mytablex.Fields("producto") = "" & mytabley.Fields("plu")
    mytablex.Fields("barras") = "" & mytabley.Fields("codi")
    mytablex.Fields("descripcio") = UCase$(Trim("" & mytabley.Fields("nom")))
    mytablex.Fields("descorto") = Mid$(Trim("" & mytabley.Fields("nom")), 1, 22)
    mytablex.Fields("presenta") = "" & mytabley.Fields("referencia")
    mytablex.Fields("familia") = Mid$("" & mytabley.Fields("familia"), 1, 6)
    mytablex.Fields("subfamilia") = "" 'Mid$("" & mytabley.Fields("subfamilia"), 1, 6)
    mytablex.Fields("seccion") = ""
    mytablex.Fields("marca") = "" 'Mid$("" & mytabley.Fields("marca"), 1, 6)
    mytablex.Fields("seccion") = "" 'Mid$("" & mytabley.Fields("ubicacio"), 1, 6)
    mytablex.Fields("categoria") = ""
    mytablex.Fields("linea") = ""
    mytablex.Fields("color") = ""
    mytablex.Fields("fabrica") = ""
    'mytablex.Fields("proveedor1") = "" & mytabley.Fields("proveedor")
    'mytablex.Fields("proveedor2") = ""
    'mytablex.Fields("codprov1") = "" & mytabley.Fields("cpa")
    'mytablex.Fields("codprov2") = ""
    mytablex.Fields("serie") = ""
    mytablex.Fields("peso") = ""
    mytablex.Fields("servicio") = ""
    mytablex.Fields("vtaund") = ""
    mytablex.Fields("oferta") = ""
    mytablex.Fields("vecaja") = "S"
    mytablex.Fields("igv") = 19
    mytablex.Fields("isc") = 0
    mytablex.Fields("pesokgr") = 0.001
    mytablex.Fields("comision") = 0
    mytablex.Fields("monedac") = "S"
    mytablex.Fields("unidad") = "UND"
    mytablex.Fields("factor") = "1"
    mytablex.Fields("costou") = Val("" & mytabley.Fields("preucost"))
    mytablex.Fields("costop") = Val("" & mytabley.Fields("preucost"))
    mytablex.Fields("monedav") = "S"
    mytablex.Fields("factor1") = "1"
    mytablex.Fields("unidad1") = "UND"
    mytablex.Fields("pventa1") = Val("" & mytabley.Fields("pv1"))

    mytablex.Fields("factor2") = "1"
    mytablex.Fields("unidad2") = "UND"
    mytablex.Fields("pventa2") = Val("" & mytabley.Fields("pv2"))

    mytablex.Fields("factor3") = "1"
    mytablex.Fields("unidad3") = "UND"
    mytablex.Fields("pventa3") = Val("" & mytabley.Fields("pv3"))

    mytablez.Seek "=", "" & mytabley.Fields("proveidor"), "" & mytabley.Fields("plu")

    If mytablez.NoMatch Then
        mytablez.AddNew
        mytablez.Fields("producto") = "" & mytabley.Fields("plu")
        mytablez.Fields("codigo") = "" & mytabley.Fields("proveidor")
        mytablez.Fields("codigop") = "" '& mytabley.Fields("cpa")

        mytablez.Update

    End If

    If Not mytablez.NoMatch Then
        mytablez.Edit
        mytablez.Fields("producto") = "" & mytabley.Fields("plu")
        mytablez.Fields("codigo") = "" & mytabley.Fields("proveidor")
        mytablez.Fields("codigop") = "" '& mytabley.Fields("cpa")
        mytablez.Update

    End If

End Sub

Sub cuentas_corrientes()

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    Dim vr

    cn.Execute ("delete from cuentac")
    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("cuentac")

    mytablex.Open "select * from cuentac ", cn, adOpenStatic, adLockOptimistic

    Do

        If mytabley.EOF Then Exit Do
        'mytablex.Seek "=", "01", "" & mytabley.Fields("tipo"), "", "" & mytabley.Fields("numero"), "" & mytabley.Fields("nro")
        'If mytablex.NoMatch Then
        mytablex.AddNew
        mytablex.Fields("local") = "02"
        mytablex.Fields("tipo") = "" & mytabley.Fields("tipo")
        mytablex.Fields("serie") = ""
        mytablex.Fields("numero") = Mid$("" & mytabley.Fields("numero"), 1, 11)
        mytablex.Fields("cuota") = "" & mytabley.Fields("nro")
        mytablex.Fields("tipoclie") = "C"
        mytablex.Fields("codigo") = "" & mytabley.Fields("codigo")
        mytablex.Fields("nombre") = busca_clientes("" & mytabley.Fields("codigo"))
   
        mytablex.Fields("fecha") = mytabley.Fields("fechavta")
        mytablex.Fields("fechav") = mytabley.Fields("fecha")
        mytablex.Fields("total") = Val("" & mytabley.Fields("valor"))
        mytablex.Fields("abono") = Val("" & mytabley.Fields("abono"))
        mytablex.Fields("interes") = Val("" & mytabley.Fields("interes"))
        mytablex.Fields("saldo") = Val("" & mytabley.Fields("saldo"))
        mytablex.Fields("estado") = "0" '& mytabley.Fields("estado")
        mytablex.Fields("vendedor") = "" & mytabley.Fields("vendedor")
        mytablex.Fields("zona") = ""
        'mytablex.Fields("dias") = Val("" & mytabley.Fields("dias"))
        mytablex.Update
        vr = DoEvents()
        xfecha = "" & mytabley.Fields("fecha")
        'End If
        'If Not mytablex.NoMatch Then
        '   mytablex.Edit
        '   mytablex.Fields("local") = "01"
        '   mytablex.Fields("tipo") = "" & mytabley.Fields("tipo")
        '   mytablex.Fields("serie") = ""
        '   mytablex.Fields("numero") = "" & mytabley.Fields("numero")
        '   mytablex.Fields("cuota") = "" & mytabley.Fields("nro")
        '   mytablex.Fields("tipoclie") = "C"
        '   mytablex.Fields("codigo") = "" & mytabley.Fields("codigo")
        'mytablex.Fields("nombre") = "" & mytabley.Fields("nombre")
        '   mytablex.Fields("fecha") = "" & mytabley.Fields("fechavta")
        '   mytablex.Fields("fechav") = mytabley.Fields("fecha")
        '   mytablex.Fields("total") = Val("" & mytabley.Fields("valor"))
        '   mytablex.Fields("abono") = Val("" & mytabley.Fields("abono"))
        '   mytablex.Fields("interes") = Val("" & mytabley.Fields("interes"))
        '   mytablex.Fields("saldo") = Val("" & mytabley.Fields("saldo"))
        '   mytablex.Fields("estado") = "0" '& mytabley.Fields("estado")
        '   mytablex.Fields("vendedor") = "" & mytabley.Fields("vendedor")
        '   mytablex.Fields("zona") = ""
   
        '   mytablex.Update
        'End If
        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    mytablex.Close

    'MsgBox "proceso Terminado", 48, "Aviso"

End Sub

Sub pafacre()

    Dim vr

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    cn.Execute ("delete from cuentacd")
    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("pafacre")

    mytablex.Open "select * from cuentacd ", cn, adOpenStatic, adLockOptimistic

    Do

        If mytabley.EOF Then Exit Do
        'mytablex.Seek "=", "01", "" & mytabley.Fields("tipo"), "", "" & mytabley.Fields("numero"), "" & mytabley.Fields("nro")
        'If mytablex.NoMatch Then
        mytablex.AddNew
        mytablex.Fields("codigo") = Trim("" & mytabley.Fields("codigo"))
        mytablex.Fields("TIPO") = Trim("" & mytabley.Fields("tipo"))
        mytablex.Fields("serie") = ""
        mytablex.Fields("numero") = Trim("" & mytabley.Fields("numero"))
        'mytablex.Fields("acu") = Trim("" & mytabley.Fields("acu"))
        mytablex.Fields("tipo1") = Trim("" & mytabley.Fields("tipor"))
        mytablex.Fields("serie1") = ""
        mytablex.Fields("numero1") = Trim("" & mytabley.Fields("numeror"))
        mytablex.Fields("cuota") = Trim("" & mytabley.Fields("cuota"))
        mytablex.Fields("moneda") = Trim("" & mytabley.Fields("moneda"))
        mytablex.Fields("total") = Val("" & mytabley.Fields("total"))
        mytablex.Fields("paga") = Val("" & mytabley.Fields("valorp"))
        mytablex.Fields("estado") = Trim("" & mytabley.Fields("estado"))
        mytablex.Fields("paridad") = 1
        mytablex.Fields("fecha") = CVDate("" & mytabley.Fields("fecha"))
        mytablex.Fields("hora") = ""
        mytablex.Fields("usuario") = Trim("" & mytabley.Fields("cobrador"))
        mytablex.Fields("local") = "02"
        mytablex.Fields("local1") = "02"
        mytablex.Fields("tipoclie") = Trim("" & mytabley.Fields("tipo_benef"))
        mytablex.Fields("caja") = ""
        mytablex.Fields("turno") = ""
   
        mytablex.Update
        vr = DoEvents()
        xfecha = "" & mytabley.Fields("fecha")
        'End If
        'If Not mytablex.NoMatch Then
        '   mytablex.Edit
        '   mytablex.Fields("local") = "01"
        '   mytablex.Fields("tipo") = "" & mytabley.Fields("tipo")
        '   mytablex.Fields("serie") = ""
        '   mytablex.Fields("numero") = "" & mytabley.Fields("numero")
        '   mytablex.Fields("cuota") = "" & mytabley.Fields("nro")
        '   mytablex.Fields("tipoclie") = "C"
        '   mytablex.Fields("codigo") = "" & mytabley.Fields("codigo")
        'mytablex.Fields("nombre") = "" & mytabley.Fields("nombre")
        '   mytablex.Fields("fecha") = "" & mytabley.Fields("fechavta")
        '   mytablex.Fields("fechav") = mytabley.Fields("fecha")
        '   mytablex.Fields("total") = Val("" & mytabley.Fields("valor"))
        '   mytablex.Fields("abono") = Val("" & mytabley.Fields("abono"))
        '   mytablex.Fields("interes") = Val("" & mytabley.Fields("interes"))
        '   mytablex.Fields("saldo") = Val("" & mytabley.Fields("saldo"))
        '   mytablex.Fields("estado") = "0" '& mytabley.Fields("estado")
        '   mytablex.Fields("vendedor") = "" & mytabley.Fields("vendedor")
        '   mytablex.Fields("zona") = ""
   
        '   mytablex.Update
        'End If
        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    mytablex.Close

    'MsgBox "proceso Terminado", 48, "Aviso"

End Sub

Sub cuentas_corrientes1()

    Dim mydby    As Database

    Dim mytablex As Table

    Dim mytabley As Table

    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("cuentap")

    Set mytablex = mydbxglo.OpenTable("cuentap")
    mytablex.Index = "cuentac"
    Do

        If mytabley.EOF Then Exit Do
        mytablex.Seek "=", "01", "" & mytabley.Fields("tipo"), "", Mid$("" & mytabley.Fields("numero"), 1, 11), "" & mytabley.Fields("nro")

        If mytablex.NoMatch Then
            mytablex.AddNew
            mytablex.Fields("local") = "02"
            mytablex.Fields("tipo") = "" & mytabley.Fields("tipo")
            mytablex.Fields("serie") = ""
            mytablex.Fields("numero") = Mid$("" & mytabley.Fields("numero"), 1, 11)
            mytablex.Fields("cuota") = "" & mytabley.Fields("nro")
            mytablex.Fields("tipoclie") = "C"
            mytablex.Fields("codigo") = "" & mytabley.Fields("codigo")

            'mytablex.Fields("nombre") = "" & mytabley.Fields("nombre")
            If IsDate("" & mytabley.Fields("fechavta")) Then
                mytablex.Fields("fecha") = "" & mytabley.Fields("fechavta")
            Else
                mytablex.Fields("fecha") = "" & mytabley.Fields("fecha")

            End If

            mytablex.Fields("fechav") = mytabley.Fields("fecha")
            mytablex.Fields("total") = Val("" & mytabley.Fields("valor"))
            mytablex.Fields("abono") = Val("" & mytabley.Fields("abono"))
            mytablex.Fields("interes") = Val("" & mytabley.Fields("interes"))
            mytablex.Fields("saldo") = Val("" & mytabley.Fields("saldo"))
            mytablex.Fields("estado") = "0" '& mytabley.Fields("estado")
            mytablex.Fields("vendedor") = "" & mytabley.Fields("vendedor")
            mytablex.Fields("zona") = ""
            'mytablex.Fields("dias") = Val("" & mytabley.Fields("dias"))
            mytablex.Update

        End If

        If Not mytablex.NoMatch Then
            mytablex.Edit
            mytablex.Fields("local") = "02"
            mytablex.Fields("tipo") = "" & mytabley.Fields("tipo")
            mytablex.Fields("serie") = ""
            mytablex.Fields("numero") = Mid$("" & mytabley.Fields("numero"), 1, 11)
            mytablex.Fields("cuota") = "" & mytabley.Fields("nro")
            mytablex.Fields("tipoclie") = "C"
            mytablex.Fields("codigo") = "" & mytabley.Fields("codigo")
            'mytablex.Fields("nombre") = "" & mytabley.Fields("nombre")
            mytablex.Fields("fecha") = "" & mytabley.Fields("fechavta")
            mytablex.Fields("fechav") = mytabley.Fields("fecha")
            mytablex.Fields("total") = Val("" & mytabley.Fields("valor"))
            mytablex.Fields("abono") = Val("" & mytabley.Fields("abono"))
            mytablex.Fields("interes") = Val("" & mytabley.Fields("interes"))
            mytablex.Fields("saldo") = Val("" & mytabley.Fields("saldo"))
            mytablex.Fields("estado") = "0" '& mytabley.Fields("estado")
            mytablex.Fields("vendedor") = "" & mytabley.Fields("vendedor")
            mytablex.Fields("zona") = ""
            'mytablex.Fields("dias") = Val("" & mytabley.Fields("dias"))
            'mytablex.Fields("dias") = Val("" & mytabley.Fields("dias"))
            mytablex.Update

        End If

        mytabley.MoveNext
    Loop
    mytablex.Close
    MsgBox "proceso Terminado", 48, "Aviso"

End Sub

Sub ngraba_producto()

    Dim mydby    As Database

    Dim mytablea As Table

    Dim mytablex As Table

    Dim mytabley As Table

    Dim mytablez As Table

    Dim mytableb As Table 'almacen

    Dim mytablec As Table 'almacen orion ant

    If procesar <> "MAXIMO" Then
        procesar = ""
        procesar.SetFocus
        Exit Sub

    End If

    'MsgBox orionv4
    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("ACCESORI")
    'Set mytablec = mydby.OpenTable("almacen")
    'mytablec.Index = "alma"

    Set mytableb = mydbxglo.OpenTable("almacen")
    mytableb.Index = "almacen"

    Set mytablex = mydbxglo.OpenTable("producto")
    mytablex.Index = "producto"
    Set mytablez = mydbxglo.OpenTable("codprov")
    mytablez.Index = "codprov"
    Set mytablea = mydbxglo.OpenTable("precios")
    mytablea.Index = "tprecios"
    Do

        If mytabley.EOF Then Exit Do
        mytablex.Seek "=", "" & mytabley.Fields("producto")

        If mytablex.NoMatch Then
            mytablex.AddNew
            npone_registro mytablex, mytabley, mytablez, mytablea, mytableb, mytablec
            mytablex.Update

        End If

        If Not mytablex.NoMatch Then
            mytablex.Edit
            npone_registro mytablex, mytabley, mytablez, mytablea, mytableb, mytablec
            mytablex.Update

        End If

        mytabley.MoveNext
    Loop
    mytablex.Close
    mytabley.Close
    mytablez.Close
    mytableb.Close
    'mytablec.Close
    mydby.Close
    MsgBox "proceso Terminado", 48, "Aviso"

End Sub

Sub npone_registro(mytablex As Table, _
                   mytabley As Table, _
                   mytablez As Table, _
                   mytablea As Table, _
                   mytableb As Table, _
                   mytablec As Table)
    mytablex.Fields("producto") = "" & mytabley.Fields("producto")
    mytablex.Fields("barras") = "" '& mytabley.Fields("barras")
    mytablex.Fields("descripcio") = Trim("" & mytabley.Fields("descrip2"))
    mytablex.Fields("descorto") = Trim(Mid$("" & mytabley.Fields("descrip2"), 1, 20))
    mytablex.Fields("oferta") = "" '& mytabley.Fields("flagremate")
    mytablex.Fields("familia") = "ACCESO" 'Mid$("" & mytabley.Fields("descripcio"), 1, 2)
    mytablex.Fields("subfamilia") = "" 'Mid$("" & mytabley.Fields("descripcio"), 1, 2)
    mytablex.Fields("seccion") = "" '& mytabley.Fields("seccion")
    mytablex.Fields("marca") = "" & Mid$("" & mytabley.Fields("marca"), 1, 6)
    mytablex.Fields("categoria") = "" '& mytabley.Fields("categoria")
    mytablex.Fields("linea") = "" '& mytabley.Fields("flagtalla")
    mytablex.Fields("color") = "" '& mytabley.Fields("color")
    mytablex.Fields("fabrica") = ""
    'mytablex.Fields("proveedor1") = "" & mytabley.Fields("proveedor")
    'mytablex.Fields("proveedor2") = ""
    'mytablex.Fields("codprov1") = "" & mytabley.Fields("cpa")
    'mytablex.Fields("codprov2") = ""
    mytablex.Fields("serie") = ""
    mytablex.Fields("peso") = "" '& mytabley.Fields("balanza")
    mytablex.Fields("servicio") = ""
    mytablex.Fields("vtaund") = "" '& mytabley.Fields("flagunidad")
    mytablex.Fields("vecaja") = "S"
    mytablex.Fields("igv") = 19 'Val("" & mytabley.Fields("igv"))
    mytablex.Fields("isc") = 0
    mytablex.Fields("pesokgr") = 0.001
    mytablex.Fields("comision") = 0
    mytablex.Fields("monedac") = "S" '& mytabley.Fields("monedac")
    mytablex.Fields("unidad") = "UND" '& mytabley.Fields("unidad")
    mytablex.Fields("factor") = 1 'Val("" & mytabley.Fields("factor"))
    mytablex.Fields("costou") = 0 'Val("" & mytabley.Fields("costopaqu"))
    mytablex.Fields("costop") = 0 'Val("" & mytabley.Fields("costopaqp"))
    mytablex.Fields("monedav") = "S" ' & mytabley.Fields("moneda")

    'mytablex.Fields("factor1") = "" & mytabley.Fields("factor1")
    'mytablex.Fields("unidad1") = "" & mytabley.Fields("unidad1")
    'mytablex.Fields("pventa1") = "" & mytabley.Fields("pventa1")

    'mytablex.Fields("factor2") = "" & mytabley.Fields("factor2")
    'mytablex.Fields("unidad2") = "" & mytabley.Fields("unidad2")
    'mytablex.Fields("pventa2") = "" & mytabley.Fields("pventa2")

    'mytablex.Fields("factor3") = "" & mytabley.Fields("factor3")
    'mytablex.Fields("unidad3") = "" & mytabley.Fields("unidad3")
    'mytablex.Fields("pventa3") = "" & mytabley.Fields("pventa3")

    'mytablex.Fields("factor4") = "" & mytabley.Fields("factor4")
    'mytablex.Fields("unidad4") = "" & mytabley.Fields("unidad4")
    'mytablex.Fields("pventa4") = "" & mytabley.Fields("pventa4")

    'mytablex.Fields("factor5") = "" & mytabley.Fields("factor5")
    'mytablex.Fields("unidad5") = "" & mytabley.Fields("unidad5")
    'mytablex.Fields("pventa5") = "" & mytabley.Fields("pventa5")

    'mytablex.Fields("factor6") = "" & mytabley.Fields("factor6")
    'mytablex.Fields("unidad6") = "" & mytabley.Fields("unidad6")
    'mytablex.Fields("pventa6") = "" & mytabley.Fields("pventa6")

    'mytablex.Fields("factor7") = "" & mytabley.Fields("factor7")
    'mytablex.Fields("unidad7") = "" & mytabley.Fields("unidad7")
    'mytablex.Fields("pventa7") = "" & mytabley.Fields("pventa7")

    'mytablex.Fields("factor8") = "" & mytabley.Fields("factor8")
    'mytablex.Fields("unidad8") = "" & mytabley.Fields("unidad8")
    'mytablex.Fields("pventa8") = Val("" & mytabley.Fields("pventa8"))

    'mytablex.Fields("factor9") = "" & mytabley.Fields("factor9")
    'mytablex.Fields("unidad9") = "" & mytabley.Fields("unidad9")
    'mytablex.Fields("pventa9") = Val("" & mytabley.Fields("pventa9"))

    'mytablex.Fields("factor10") = "" & mytabley.Fields("factor10")
    'mytablex.Fields("unidad10") = "" & mytabley.Fields("unidad10")
    'mytablex.Fields("pventa10") = Val("" & mytabley.Fields("pventa10"))
    'mytablex.Fields("SECcion") = "" & mytabley.Fields("seccion")

    'If "" & mytabley.Fields("seccion") = "1" Or "" & mytabley.Fields("seccion") = "3" Then
    '   mytablex.Fields("c11") = "1"
    'End If
    'If "" & mytabley.Fields("seccion") = "2" Or "" & mytabley.Fields("seccion") = "6" Then
    '   mytablex.Fields("c12") = "1"
    'End If
    'If "" & mytabley.Fields("seccion") = "4" Then
    '   mytablex.Fields("c13") = "1"
    'End If
    'If "" & mytabley.Fields("seccion") = "5" Then
    '   mytablex.Fields("c14") = "1"
    'End If
    'mytablez.Seek "=", "" & mytabley.Fields("proveedor"), "" & mytabley.Fields("producto")
    'If mytablez.NoMatch Then
    '   mytablez.AddNew
    '   npone_detalle mytablez, mytabley
    '   mytablez.Update
    'End If
    'If Not mytablez.NoMatch Then
    '   mytablez.Edit
    '   npone_detalle mytablez, mytabley
    '   mytablez.Update
    'End If
    'grabando precios al local nro 1
    mytablea.Seek "=", "" & mytabley.Fields("producto"), "01"

    If mytablea.NoMatch Then
        mytablea.AddNew
        npone_detalle01 mytablea, mytabley
        mytablea.Update

    End If

    If Not mytablea.NoMatch Then
        mytablea.Edit
        npone_detalle01 mytablea, mytabley
        mytablea.Update

    End If

    'grabando almacen

    'mytablec.Seek "=", "" & mytabley.Fields("producto"), "01"
    'If Not mytablec.NoMatch Then
    'mytableb.Seek "=", "01", "" & mytablec.Fields("producto"), almacen
    'If mytableb.NoMatch Then
    '   mytableb.AddNew
    '   mytableb.Fields("local") = "01"
    '   mytableb.Fields("producto") = "" & mytablec.Fields("producto")
    '   mytableb.Fields("bodega") = almacen
    '   mytableb.Fields("saldo") = Val("" & mytablec.Fields("saldo"))
    '   mytableb.Update
    'End If
    'If Not mytableb.NoMatch Then
    '   mytableb.Edit
    '   mytableb.Fields("local") = "01"
    '   mytableb.Fields("producto") = "" & mytablec.Fields("producto")
    '   mytableb.Fields("bodega") = almacen
    '   mytableb.Fields("saldo") = Val("" & mytablec.Fields("saldo"))
    '   mytableb.Update
    'End If
    'End If

End Sub

Sub npone_detalle01(mytablex As Table, mytabley As Table)
    mytablex.Fields("local") = "02"
    mytablex.Fields("producto") = "" & mytabley.Fields("producto")
    mytablex.Fields("ccosto") = "" '& mytabley.Fields("seccion")
    'mytablex.Fields("monedav") = "" & mytabley.Fields("moneda")
    mytablex.Fields("factor1") = 1 'Val("" & mytabley.Fields("factor1"))
    mytablex.Fields("unidad1") = "UND" '& mytabley.Fields("unidad1")
    mytablex.Fields("pventa1") = Val("" & mytabley.Fields("p1"))

    mytablex.Fields("factor2") = 1 'Val("" & mytabley.Fields("factor2"))
    mytablex.Fields("unidad2") = "UND" '& mytabley.Fields("unidad2")
    mytablex.Fields("pventa2") = Val("" & mytabley.Fields("p2"))

    'mytablex.Fields("factor3") = 1 ' Val("" & mytabley.Fields("factor3"))
    'mytablex.Fields("unidad3") = "UND" '& mytabley.Fields("unidad3")
    'mytablex.Fields("pventa3") = Val("" & mytabley.Fields("p3"))

    'mytablex.Fields("factor4") = 1 ' Val("" & mytabley.Fields("factor4"))
    'mytablex.Fields("unidad4") = "UND" '& mytabley.Fields("unidad4")
    'mytablex.Fields("pventa4") = Val("" & mytabley.Fields("p4"))

    'mytablex.Fields("factor5") = 12 'Val("" & mytabley.Fields("factor5"))
    'mytablex.Fields("unidad5") = "DOC" '& mytabley.Fields("unidad5")
    'mytablex.Fields("pventa5") = Val("" & mytabley.Fields("p5"))

    'mytablex.Fields("factor6") = 12 'Val("" & mytabley.Fields("factor6"))
    'mytablex.Fields("unidad6") = "DOC" '& mytabley.Fields("unidad6")
    'mytablex.Fields("pventa6") = Val("" & mytabley.Fields("p6"))

    'mytablex.Fields("factor7") = 12 'Val("" & mytabley.Fields("factor7"))
    'mytablex.Fields("unidad7") = "DOC" '& mytabley.Fields("unidad7")
    'mytablex.Fields("pventa7") = Val("" & mytabley.Fields("p7"))

    'mytablex.Fields("factor8") = Val("" & mytabley.Fields("factor8"))
    'mytablex.Fields("unidad8") = "" & mytabley.Fields("unidad8")
    'mytablex.Fields("pventa8") = Val("" & mytabley.Fields("pventa8"))

    'mytablex.Fields("factor9") = Val("" & mytabley.Fields("factor9"))
    'mytablex.Fields("unidad9") = "" & mytabley.Fields("unidad9")
    'mytablex.Fields("pventa9") = Val("" & mytabley.Fields("pventa9"))

    'mytablex.Fields("factor10") = Val("" & mytabley.Fields("factor10"))
    'mytablex.Fields("unidad10") = "" & mytabley.Fields("unidad10")
    'mytablex.Fields("pventa10") = Val("" & mytabley.Fields("pventa10"))

End Sub

Sub npone_detalle(mytablex As Table, rs As Table)
    mytablex.Fields("producto") = "" & rs.Fields("producto")
    mytablex.Fields("codigo") = "" & rs.Fields("proveedor")
    mytablex.Fields("codigop") = "" & rs.Fields("cpa")

End Sub

Sub graba_xxfamiliasx()

    Dim mydby    As Database

    Dim mytablex As Table

    Dim mytablez As Table

    Dim mytablea As Table

    Dim mytableb As Table

    Dim mytablec As Table

    Dim mytabled As Table

    Dim mytabley As Snapshot

    Dim buf      As String

    Dim vr

    Dim ind As Long

    'Set mydby = OpenDatabase("c:\URANIO\X", False, False, "foxpro 2.5;")
    Set mytabley = mydbxglo.CreateSnapshot("producto")

    Set mytablez = mydbxglo.OpenTable("familia")
    mytablez.Index = "familia"
    Set mytableb = mydbxglo.OpenTable("marca")
    mytableb.Index = "marca"

    Do

        If mytabley.EOF Then Exit Do
        'graba familias
        mytablez.Seek "=", Mid$("" & mytabley.Fields("familia"), 1, 6)

        If mytablez.NoMatch Then
            mytablez.AddNew
            mytablez.Fields("familia") = Mid$("" & mytabley.Fields("familia"), 1, 6)
            mytablez.Fields("descripcio") = Mid$("" & mytabley.Fields("familia"), 1, 15)
            mytablez.Update

        End If

        If Not mytablez.NoMatch Then
            mytablez.Edit
            mytablez.Fields("familia") = Mid$("" & mytabley.Fields("familia"), 1, 6)
            mytablez.Fields("descripcio") = Mid$("" & mytabley.Fields("familia"), 1, 15)
            mytablez.Update

        End If

        'graba seccion
        mytableb.Seek "=", Mid$("" & mytabley.Fields("marca"), 1, 6)

        If mytableb.NoMatch Then
            mytableb.AddNew
            mytableb.Fields("marca") = Mid$("" & mytabley.Fields("marca"), 1, 6)
            mytableb.Fields("descripcio") = Mid$("" & mytabley.Fields("marca"), 1, 15)
            mytableb.Update

        End If

        If Not mytableb.NoMatch Then
            mytableb.Edit
            mytableb.Fields("marca") = Mid$("" & mytabley.Fields("marca"), 1, 6)
            mytableb.Fields("descripcio") = Mid$("" & mytabley.Fields("marca"), 1, 15)
            mytableb.Update

        End If

        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    mytabley.Close

    MsgBox "proceso Terminado", 48, "Aviso"

End Sub

Function pasar_datauno()

    Dim mytablex As Table

    Dim mytabley As Table

    Dim mytablea As Table

    Dim mydby    As Database

    Dim sdx      As Double

    Dim I        As Integer

    Dim sdx1     As String

    Dim vr

    Dim c As Integer

    Set mydby = OpenDatabase("g:\nuevo_erp\excell", False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("montura1")

    Set mytablex = mydbxglo.OpenTable("producto")
    mytablex.Index = "producto"
    Set mytablea = mydbxglo.OpenTable("precios")
    mytablea.Index = "tprecios"
    sdx = 0
    Do
        c = 1

        'MsgBox "" & mytabley.Fields(0)
        'End
        If mytabley.EOF Then Exit Do

        For I = 10 To 43

            If Len("" & mytabley.Fields(I)) > 0 Then
                sdx = sdx + 1
                sdx1 = "MT" & Format(sdx, "0000")
                mytablex.Seek "=", sdx1

                If mytablex.NoMatch Then
                    mytablex.AddNew
                    graba_productoxs mytablex, mytabley, sdx1, I, mytablea
                    mytablex.Update

                End If

            End If

            dd = "" & I & " " & mytabley.Fields(I)
            vr = DoEvents()
        Next I

        'Exit Function
        mytabley.MoveNext
    Loop
    mytablex.Close
    mytabley.Close

End Function

Sub graba_productoxs(mytablex As Table, _
                     mytabley As Table, _
                     sdx1 As String, _
                     I As Integer, _
                     mytablea As Table)
    mytablex.Fields("producto") = sdx1
    mytablex.Fields("barras") = "" '& mytabley.Fields("barras")
    mytablex.Fields("descripcio") = Trim("" & mytabley.Fields("nom")) + " " + Trim("" & mytabley.Fields("marca")) + " " + Trim("" & mytabley.Fields(I))
    mytablex.Fields("descorto") = Mid$(Trim(Trim("" & mytabley.Fields("nom")) + "" + Trim("" & mytabley.Fields("marca"))), 1, 20)
    mytablex.Fields("oferta") = "" '& mytabley.Fields("flagremate")
    mytablex.Fields("familia") = "MONTUR" 'Mid$("" & mytabley.Fields("descripcio"), 1, 2)
    mytablex.Fields("subfamilia") = "" 'Mid$("" & mytabley.Fields("descripcio"), 1, 2)
    mytablex.Fields("seccion") = "" '& mytabley.Fields("seccion")
    mytablex.Fields("marca") = "" & Mid$("" & mytabley.Fields("marca"), 1, 6)
    mytablex.Fields("categoria") = "" '& mytabley.Fields("categoria")
    mytablex.Fields("linea") = "" '& mytabley.Fields("flagtalla")
    mytablex.Fields("color") = "" '& mytabley.Fields("color")
    mytablex.Fields("fabrica") = ""
    'mytablex.Fields("proveedor1") = "" & mytabley.Fields("proveedor")
    'mytablex.Fields("proveedor2") = ""
    'mytablex.Fields("codprov1") = "" & mytabley.Fields("cpa")
    'mytablex.Fields("codprov2") = ""
    mytablex.Fields("serie") = ""
    mytablex.Fields("peso") = "" '& mytabley.Fields("balanza")
    mytablex.Fields("servicio") = ""
    mytablex.Fields("vtaund") = "" '& mytabley.Fields("flagunidad")
    mytablex.Fields("vecaja") = "S"
    mytablex.Fields("igv") = 19 'Val("" & mytabley.Fields("igv"))
    mytablex.Fields("isc") = 0
    mytablex.Fields("pesokgr") = 0.001
    mytablex.Fields("comision") = 0
    mytablex.Fields("monedac") = "S" '& mytabley.Fields("monedac")
    mytablex.Fields("unidad") = "UND" '& mytabley.Fields("unidad")
    mytablex.Fields("factor") = 1 'Val("" & mytabley.Fields("factor"))
    mytablex.Fields("costou") = 0 'Val("" & mytabley.Fields("costopaqu"))
    mytablex.Fields("costop") = 0 'Val("" & mytabley.Fields("costopaqp"))
    mytablex.Fields("monedav") = "S" ' & mytabley.Fields("moneda")

    mytablea.Seek "=", sdx1, "01"

    If mytablea.NoMatch Then
        mytablea.AddNew
        npone_detalle02 mytablea, mytabley, sdx1
        mytablea.Update

    End If

    If Not mytablea.NoMatch Then
        mytablea.Edit
        npone_detalle02 mytablea, mytabley, sdx1
        mytablea.Update

    End If

End Sub

Sub npone_detalle02(mytablex As Table, mytabley As Table, sdx)
    mytablex.Fields("local") = "02"
    mytablex.Fields("producto") = sdx
    mytablex.Fields("ccosto") = "" '& mytabley.Fields("seccion")
    'mytablex.Fields("monedav") = "" & mytabley.Fields("moneda")
    mytablex.Fields("factor1") = 1 'Val("" & mytabley.Fields("factor1"))
    mytablex.Fields("unidad1") = "UND" '& mytabley.Fields("unidad1")
    mytablex.Fields("pventa1") = Val("" & mytabley.Fields("p1"))

    mytablex.Fields("factor2") = 1 'Val("" & mytabley.Fields("factor2"))
    mytablex.Fields("unidad2") = "UND" '& mytabley.Fields("unidad2")
    mytablex.Fields("pventa2") = Val("" & mytabley.Fields("p2"))

    mytablex.Fields("factor3") = 1 ' Val("" & mytabley.Fields("factor3"))
    mytablex.Fields("unidad3") = "UND" '& mytabley.Fields("unidad3")
    mytablex.Fields("pventa3") = Val("" & mytabley.Fields("p3"))

    mytablex.Fields("factor4") = 1 ' Val("" & mytabley.Fields("factor4"))
    mytablex.Fields("unidad4") = "UND" '& mytabley.Fields("unidad4")
    mytablex.Fields("pventa4") = Val("" & mytabley.Fields("p4"))

    mytablex.Fields("factor5") = 12 'Val("" & mytabley.Fields("factor5"))
    mytablex.Fields("unidad5") = "DOC" '& mytabley.Fields("unidad5")
    mytablex.Fields("pventa5") = Val("" & mytabley.Fields("p5"))

    mytablex.Fields("factor6") = 12 'Val("" & mytabley.Fields("factor6"))
    mytablex.Fields("unidad6") = "DOC" '& mytabley.Fields("unidad6")
    mytablex.Fields("pventa6") = Val("" & mytabley.Fields("p6"))

    'mytablex.Fields("factor7") = 12 'Val("" & mytabley.Fields("factor7"))
    'mytablex.Fields("unidad7") = "DOC" '& mytabley.Fields("unidad7")
    'mytablex.Fields("pventa7") = Val("" & mytabley.Fields("p7"))

End Sub

Sub graba_producto_guido()

    Dim mydby    As Database

    Dim mytablea As Table

    Dim mytablex As Table

    Dim mytabley As Table

    Dim mytableb As Table

    If procesar <> "Guido" Then
        procesar = ""
        procesar.SetFocus
        Exit Sub

    End If

    'MsgBox orionv4
    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("PRECIOS")
    Set mytablex = mydbxglo.OpenTable("producto")
    mytablex.Index = "producto"
    Set mytablea = mydbxglo.OpenTable("precios")
    mytablea.Index = "tprecios"
    Set mytableb = mydbxglo.OpenTable("familia")
    mytableb.Index = "familia"

    Do

        If mytabley.EOF Then Exit Do
        mytablex.Seek "=", "" & mytabley.Fields("producto")

        If mytablex.NoMatch Then

            'mytablex.AddNew
            'pone_registro_guido mytablex, mytabley, mytablea
            'mytablex.Update
        End If

        If Not mytablex.NoMatch Then

            'mytablex.Edit
            'pone_registro_guido mytablex, mytabley, mytablea
            'mytablex.Update
        End If

        'FAMILIAS
        mytableb.Seek "=", Mid$("" & mytabley.Fields("familia"), 1, 6)

        If mytableb.NoMatch Then
            mytableb.AddNew
            mytableb.Fields("familia") = Mid$("" & mytabley.Fields("familia"), 1, 6)
            mytableb.Fields("descripcio") = "" & mytabley.Fields("desfam")
            mytableb.Update

        End If

        If Not mytableb.NoMatch Then
            mytableb.Edit
            mytableb.Fields("familia") = Mid$("" & mytabley.Fields("familia"), 1, 6)
            mytableb.Fields("descripcio") = "" & mytabley.Fields("desfam")
            mytableb.Update

        End If

        mytabley.MoveNext
    Loop
    mytablex.Close
    mytabley.Close
    mytablea.Close
    mydby.Close
    MsgBox "proceso Terminado", 48, "Aviso"

End Sub

Sub pone_registro_guido(mytablex As Table, mytabley As Table, mytablea As Table)
    mytablex.Fields("producto") = "" & mytabley.Fields("producto")
    mytablex.Fields("barras") = ""
    mytablex.Fields("descripcio") = UCase("" & mytabley.Fields("descripcio"))
    mytablex.Fields("descorto") = UCase(Mid$("" & mytabley.Fields("descripcio"), 1, 20))
    'mytablex.Fields("oferta") = "" & mytabley.Fields("flagremate")
    mytablex.Fields("familia") = Mid$("" & mytabley.Fields("familia"), 1, 6)
    mytablex.Fields("subfamilia") = ""
    mytablex.Fields("seccion") = ""
    mytablex.Fields("marca") = ""
    mytablex.Fields("categoria") = ""
    'mytablex.Fields("linea") = "" & mytabley.Fields("flagtalla")
    mytablex.Fields("color") = ""
    mytablex.Fields("fabrica") = ""
    'mytablex.Fields("proveedor1") = "" & mytabley.Fields("proveedor")
    'mytablex.Fields("proveedor2") = ""
    'mytablex.Fields("codprov1") = "" & mytabley.Fields("cpa")
    'mytablex.Fields("codprov2") = ""
    mytablex.Fields("serie") = ""
    mytablex.Fields("peso") = ""
    mytablex.Fields("servicio") = ""
    'mytablex.Fields("vtaund") = "" & mytabley.Fields("flagunidad")
    mytablex.Fields("vecaja") = "S"
    mytablex.Fields("igv") = 19
    mytablex.Fields("isc") = 0
    mytablex.Fields("pesokgr") = 0.001
    mytablex.Fields("comision") = 0
    mytablex.Fields("monedac") = "S"
    mytablex.Fields("unidad") = "UND"
    mytablex.Fields("factor") = 1
    mytablex.Fields("costou") = 0
    mytablex.Fields("costop") = 0
    mytablex.Fields("monedav") = "S"

    'mytablex.Fields("factor1") = "" & mytabley.Fields("factor1")
    'mytablex.Fields("unidad1") = "" & mytabley.Fields("unidad1")
    'mytablex.Fields("pventa1") = "" & mytabley.Fields("pventa1")

    'mytablex.Fields("factor2") = "" & mytabley.Fields("factor2")
    'mytablex.Fields("unidad2") = "" & mytabley.Fields("unidad2")
    'mytablex.Fields("pventa2") = "" & mytabley.Fields("pventa2")

    'mytablex.Fields("factor3") = "" & mytabley.Fields("factor3")
    'mytablex.Fields("unidad3") = "" & mytabley.Fields("unidad3")
    'mytablex.Fields("pventa3") = "" & mytabley.Fields("pventa3")

    'mytablex.Fields("factor4") = "" & mytabley.Fields("factor4")
    'mytablex.Fields("unidad4") = "" & mytabley.Fields("unidad4")
    'mytablex.Fields("pventa4") = "" & mytabley.Fields("pventa4")

    'mytablex.Fields("factor5") = "" & mytabley.Fields("factor5")
    'mytablex.Fields("unidad5") = "" & mytabley.Fields("unidad5")
    'mytablex.Fields("pventa5") = "" & mytabley.Fields("pventa5")

    'mytablex.Fields("factor6") = "" & mytabley.Fields("factor6")
    'mytablex.Fields("unidad6") = "" & mytabley.Fields("unidad6")
    'mytablex.Fields("pventa6") = "" & mytabley.Fields("pventa6")

    'mytablex.Fields("factor7") = "" & mytabley.Fields("factor7")
    'mytablex.Fields("unidad7") = "" & mytabley.Fields("unidad7")
    'mytablex.Fields("pventa7") = "" & mytabley.Fields("pventa7")

    'mytablex.Fields("factor8") = "" & mytabley.Fields("factor8")
    'mytablex.Fields("unidad8") = "" & mytabley.Fields("unidad8")
    'mytablex.Fields("pventa8") = Val("" & mytabley.Fields("pventa8"))

    'mytablex.Fields("factor9") = "" & mytabley.Fields("factor9")
    'mytablex.Fields("unidad9") = "" & mytabley.Fields("unidad9")
    'mytablex.Fields("pventa9") = Val("" & mytabley.Fields("pventa9"))

    'mytablex.Fields("factor10") = "" & mytabley.Fields("factor10")
    'mytablex.Fields("unidad10") = "" & mytabley.Fields("unidad10")
    'mytablex.Fields("pventa10") = Val("" & mytabley.Fields("pventa10"))
    'mytablex.Fields("SECcion") = "" & mytabley.Fields("seccion")

    'If "" & mytabley.Fields("seccion") = "1" Or "" & mytabley.Fields("seccion") = "3" Then
    '   mytablex.Fields("c11") = "1"
    'End If
    'If "" & mytabley.Fields("seccion") = "2" Or "" & mytabley.Fields("seccion") = "6" Then
    '   mytablex.Fields("c12") = "1"
    'End If
    'If "" & mytabley.Fields("seccion") = "4" Then
    '   mytablex.Fields("c13") = "1"
    'End If
    'If "" & mytabley.Fields("seccion") = "5" Then
    '   mytablex.Fields("c14") = "1"
    'End If
    'grabando precios al local nro 1

    mytablea.Seek "=", "" & mytabley.Fields("producto"), "01"

    If mytablea.NoMatch Then
        mytablea.AddNew
        pone_detalle01_guido mytablea, mytabley
        mytablea.Update

    End If

    If Not mytablea.NoMatch Then
        mytablea.Edit
        pone_detalle01_guido mytablea, mytabley
        mytablea.Update

    End If

End Sub

Sub pone_detalle01_guido(mytablex As Table, mytabley As Table)
    mytablex.Fields("local") = "02"
    mytablex.Fields("producto") = "" & mytabley.Fields("producto")
    mytablex.Fields("ccosto") = "" '& mytabley.Fields("seccion")
    'mytablex.Fields("monedav") = "" & mytabley.Fields("moneda")
    mytablex.Fields("factor1") = 1 'Val("" & mytabley.Fields("factor1"))
    mytablex.Fields("unidad1") = "UND"  '"" & mytabley.Fields("unidad1")
    mytablex.Fields("pventa1") = Val("" & mytabley.Fields("pventa"))

End Sub

Sub graba_ingresos()

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Snapshot

    Dim buf      As String

    Dim vr

    '--------------eliminando----------------------------------------
    cn.Execute "DELETE FROM recibo "
    xfecha = ""
    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.CreateSnapshot("select * from ingreso")
    mytablex.Open "select * from recibo ", cn, adOpenStatic, adLockOptimistic

    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew
        pone_registro_ingreso mytablex, mytabley
        mytablex.Update
        vr = DoEvents
        xfecha = "" & mytabley.Fields("fecha")
        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    mytabley.Close
    mytablex.Close
    'MsgBox "Recibo-proceso Terminado", 48, "Aviso"

End Sub

Sub pone_registro_ingreso(mytablez As ADODB.Recordset, mytabler As Snapshot)
    mytablez.Fields("estado") = "" & mytabler.Fields("estado")
    mytablez.Fields("local") = "02" '& mytabler.Fields("local")
    mytablez.Fields("tipo") = "" & mytabler.Fields("tipo")
    mytablez.Fields("serie") = "" '& mytabler.Fields("serie")
    mytablez.Fields("numero") = Mid$("" & mytabler.Fields("numero"), 1, 11)
    mytablez.Fields("tipoclie") = "C"

    If "" & mytabler.Fields("acu") = "X" Then
        mytablez.Fields("acu") = "W"
        mytablez.Fields("afecta") = "P"

    End If

    If "" & mytabler.Fields("acu") = "Y" Then
        mytablez.Fields("acu") = "V"
        mytablez.Fields("afecta") = "C"

    End If

    mytablez.Fields("codigo") = "" & mytabler.Fields("codigo")
    mytablez.Fields("codigo") = "" & mytabler.Fields("codigo")
    mytablez.Fields("nombre") = "" & mytabler.Fields("nombre")
    mytablez.Fields("numero") = Mid$("" & mytabler.Fields("numero"), 1, 11)
    mytablez.Fields("fecha") = "" & mytabler.Fields("fecha")
    mytablez.Fields("moneda") = "" & mytabler.Fields("moneda")
    mytablez.Fields("total") = Val("" & mytabler.Fields("total"))
    mytablez.Fields("caja") = "" & mytabler.Fields("caja")
    mytablez.Fields("turno") = "" & mytabler.Fields("turno")
    mytablez.Fields("usuario") = "" & mytabler.Fields("usuario")
    mytablez.Fields("observa") = "" 'Mid$("" & mytabler.Fields("observa"), 1, 15)
    mytablez.Fields("estado") = "" & mytabler.Fields("estado")

End Sub

Sub graba_producto_chiclayo()

    Dim mydby    As Database

    Dim mytablea As Table

    Dim mytablex As Table

    Dim mytabley As Table

    Dim mytableb As Table

    Dim mytablec As Table

    If procesar <> "Chiclayo" Then
        procesar = ""
        procesar.SetFocus
        Exit Sub

    End If

    'MsgBox orionv4
    Set mydby = OpenDatabase("\", False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("nova")
    Set mytablex = mydbxglo.OpenTable("producto")
    mytablex.Index = "producto"
    Set mytablea = mydbxglo.OpenTable("precios")
    mytablea.Index = "tprecios"
    Set mytableb = mydbxglo.OpenTable("familia")
    mytableb.Index = "familia"
    Set mytablec = mydbxglo.OpenTable("saldoini")
    mytablec.Index = "saldoini"

    Do

        If mytabley.EOF Then Exit Do
        mytablex.Seek "=", "" & mytabley.Fields("producto")

        If mytablex.NoMatch Then
            mytablex.AddNew
            pone_registro_chiclayo mytablex, mytabley, mytablea
            mytablex.Update

        End If

        If Not mytablex.NoMatch Then
            mytablex.Edit
            pone_registro_chiclayo mytablex, mytabley, mytablea
            mytablex.Update

        End If

        'FAMILIAS
        mytableb.Seek "=", Mid$("" & mytabley.Fields("familia"), 1, 6)

        If mytableb.NoMatch Then
            mytableb.AddNew
            mytableb.Fields("familia") = Mid$("" & mytabley.Fields("familia"), 1, 6)
            mytableb.Fields("descripcio") = Mid$("" & mytabley.Fields("desfam"), 1, 15)
            mytableb.Update

        End If

        If Not mytableb.NoMatch Then
            mytableb.Edit
            mytableb.Fields("familia") = Mid$("" & mytabley.Fields("familia"), 1, 6)
            mytableb.Fields("descripcio") = Mid$("" & mytabley.Fields("desfam"), 1, 15)
            mytableb.Update

        End If

        'saldo inicial
        'saldoinicial
        mytablec.Seek "=", "01", "" & mytabley.Fields("producto"), "01", Format(Now, "dd/mm/yyyy")

        If mytablec.NoMatch Then
            mytablec.AddNew
            mytablec.Fields("familia") = Mid$("" & mytabley.Fields("familia"), 1, 6)
            mytablec.Fields("producto") = "" & mytabley.Fields("producto")
            mytablec.Fields("DESCRIPCIO") = UCase(Mid$("" & mytabley.Fields("familia"), 1, 6)) + " " + Trim(Mid$(UCase("" & mytabley.Fields("descripcio")), 1, 56))
            mytablec.Fields("unidad") = "UND"
            mytablec.Fields("factor") = 1
            mytablec.Fields("bodega") = "01"
            mytablec.Fields("local") = "01"
            mytablec.Fields("fecha") = Format(Now, "dd/mm/yyyy")
            mytablec.Fields("cantidad") = Val("" & mytabley.Fields("stock"))
            mytablec.Update

        End If

        If Not mytablec.NoMatch Then
            mytablec.Edit
            mytablec.Fields("familia") = Mid$("" & mytabley.Fields("familia"), 1, 6)
            mytablec.Fields("producto") = "" & mytabley.Fields("producto")
            mytablec.Fields("DESCRIPCIO") = Trim("" & mytabley.Fields("descripcio"))
            mytablec.Fields("unidad") = "UND"
            mytablec.Fields("factor") = 1
            mytablec.Fields("fecha") = Format(Now, "dd/mm/yyyy")
            mytablec.Fields("bodega") = "01"
            mytablec.Fields("local") = "01"
            mytablec.Fields("cantidad") = Val("" & mytabley.Fields("stock"))
            mytablec.Update

        End If

        mytabley.MoveNext
    Loop

    mytablex.Close
    mytabley.Close
    mytablea.Close
    mytablec.Close
    mydby.Close
    MsgBox "proceso Terminado", 48, "Aviso"

End Sub

Sub pone_registro_chiclayo(mytablex As Table, mytabley As Table, mytablea As Table)
    mytablex.Fields("producto") = "" & mytabley.Fields("producto")
    mytablex.Fields("barras") = ""
    mytablex.Fields("descripcio") = UCase(Mid$("" & mytabley.Fields("familia"), 1, 6)) + " " + Trim(Mid$(UCase("" & mytabley.Fields("descripcio")), 1, 56))
    mytablex.Fields("descorto") = UCase(Mid$("" & mytabley.Fields("descripcio"), 1, 20))
    'mytablex.Fields("oferta") = "" & mytabley.Fields("flagremate")
    mytablex.Fields("familia") = Mid$("" & mytabley.Fields("familia"), 1, 6)
    mytablex.Fields("subfamilia") = ""
    mytablex.Fields("seccion") = ""
    mytablex.Fields("marca") = ""
    mytablex.Fields("categoria") = ""
    'mytablex.Fields("linea") = "" & mytabley.Fields("flagtalla")
    mytablex.Fields("color") = ""
    mytablex.Fields("fabrica") = ""
    'mytablex.Fields("proveedor1") = "" & mytabley.Fields("proveedor")
    'mytablex.Fields("proveedor2") = ""
    'mytablex.Fields("codprov1") = "" & mytabley.Fields("cpa")
    'mytablex.Fields("codprov2") = ""
    mytablex.Fields("serie") = ""
    mytablex.Fields("peso") = ""
    mytablex.Fields("servicio") = ""
    'mytablex.Fields("vtaund") = "" & mytabley.Fields("flagunidad")
    mytablex.Fields("vecaja") = "S"
    mytablex.Fields("igv") = 19
    mytablex.Fields("isc") = 0
    mytablex.Fields("pesokgr") = 0.001
    mytablex.Fields("comision") = 0
    mytablex.Fields("monedac") = "S"
    mytablex.Fields("unidad") = "UND"
    mytablex.Fields("factor") = 1
    mytablex.Fields("costou") = 0
    mytablex.Fields("costop") = 0
    mytablex.Fields("monedav") = "S"

    If Val("" & mytabley.Fields("soles")) > 0 Then
        mytablex.Fields("monedav") = "S"

    End If

    If Val("" & mytabley.Fields("dolares")) > 0 Then
        mytablex.Fields("monedav") = "D"

    End If

    mytablea.Seek "=", "" & mytabley.Fields("producto"), "01"

    If mytablea.NoMatch Then
        mytablea.AddNew
        pone_detalle01_chiclayo mytablea, mytabley
        mytablea.Update

    End If

    If Not mytablea.NoMatch Then
        mytablea.Edit
        pone_detalle01_chiclayo mytablea, mytabley
        mytablea.Update

    End If

End Sub

Sub pone_detalle01_chiclayo(mytablex As Table, mytabley As Table)
    mytablex.Fields("local") = "01"
    mytablex.Fields("producto") = "" & mytabley.Fields("producto")
    mytablex.Fields("ccosto") = "" '& mytabley.Fields("seccion")
    mytablex.Fields("factor1") = 1 'Val("" & mytabley.Fields("factor1"))
    mytablex.Fields("unidad1") = "UND"  '"" & mytabley.Fields("unidad1")
    mytablex.Fields("pventa1") = Val("" & mytabley.Fields("soles"))

    If Val("" & mytabley.Fields("soles")) > 0 Then
        mytablex.Fields("pventa1") = Val("" & mytabley.Fields("soles"))

    End If

    If Val("" & mytabley.Fields("dolares")) > 0 Then
        mytablex.Fields("pventa1") = Val("" & mytabley.Fields("dolares"))

    End If

End Sub

Sub graba_equiva()

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    If procesar <> "Procesar" Then
        procesar = ""
        procesar.SetFocus
        Exit Sub

    End If

    cn.Execute ("delete from productb")
    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("EQUIVA")

    'mytablex.Index = "marca"

    mytablex.Open "select * from productb ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew
        mytablex.Fields("producto") = "" & mytabley.Fields("producto")
        mytablex.Fields("barras") = "" & mytabley.Fields("barras")
        'mytablex.Fields("local") = "" & mytabley.Fields("local")
        mytablex.Update
        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    mytablex.Close
    MsgBox "Equiva proceso Terminado", 48, "Aviso"

End Sub

Function busca_clientes(buf As String) As String

    Dim mytablex As Table

    Set mytablex = mydbxglo.OpenTable("clientes")
    mytablex.Index = "codigo"
    mytablex.Seek "=", buf

    If Not mytablex.NoMatch Then
        busca_clientes = "" & mytablex.Fields("nombre")

    End If

    mytablex.Close

End Function

Sub cajamarca_producto()

    Dim I   As Integer

    Dim j   As Integer

    Dim P1  As Double

    Dim u1  As String

    Dim f1  As Double

    Dim ind As Long

    ReDim buferx(10) As String
    ReDim buferx1(10) As Double

    Dim mydby    As Database

    Dim mytablea As Table

    Dim mytablex As Table

    Dim mytabley As Table

    Dim mytablez As Table

    Dim mytableb As Table 'almacen

    Dim mytablec As Table 'almacen orion ant

    Dim mytabled As Table 'familia

    Dim vr

    Dim mytablef As Table 'lista precios

    Dim mytableh As Table

    If procesar <> "CAJAMARCA" Then
        procesar = ""
        procesar.SetFocus
        Exit Sub

    End If

    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")  'cajamarca
    Set mytabley = mydby.OpenTable("producto")

    Set mytablef = mydby.OpenTable("precio")  'cajamarca
    mytablef.Index = "precio"

    Set mytabled = mydbxglo.OpenTable("familia")
    mytabled.Index = "familia"

    Set mytableb = mydbxglo.OpenTable("almacen")
    mytableb.Index = "almacen"

    Set mytablex = mydbxglo.OpenTable("producto")
    mytablex.Index = "producto"

    Set mytablea = mydbxglo.OpenTable("precios")
    mytablea.Index = "tprecios"

    Set mytableh = mydbxglo.OpenTable("marca")
    mytableh.Index = "marca"

    Do

        If mytabley.EOF Then Exit Do

        'GRABANDO FAMILIA
        mytablef.Seek "=", Mid$(Trim("" & mytabley.Fields("codigo")), 1, 15)

        If Not mytablef.NoMatch Then
            mytabled.Seek "=", Mid$(Trim("" & mytablef.Fields("catEgoria")), 1, 6)

            If mytabled.NoMatch Then
                mytabled.AddNew
                mytabled.Fields("familia") = Mid$(Trim("" & mytablef.Fields("categoria")), 1, 6)
                mytabled.Fields("descripcio") = Mid$(Trim("" & mytablef.Fields("categoria")), 1, 6)
                mytabled.Update

            End If

        End If

        'grabando marca
        mytableh.Seek "=", Mid$(Trim("" & mytabley.Fields("marca")), 1, 6)

        If mytableh.NoMatch Then
            mytableh.AddNew
            mytableh.Fields("marca") = Mid$(Trim("" & mytabley.Fields("marca")), 1, 6)
            mytableh.Fields("descripcio") = Mid$(Trim("" & mytabley.Fields("marca")), 1, 15)
            mytableh.Update

        End If

        'grabando producto
        mytablex.Seek "=", Mid$(Trim("" & mytabley.Fields("codigo")), 1, 15)

        If mytablex.NoMatch Then
            mytablex.AddNew
            mytablef.Seek "=", Mid$(Trim("" & mytabley.Fields("codigo")), 1, 15)

            If Not mytablef.NoMatch Then
                mytablex.Fields("familia") = Mid$(Trim("" & mytablef.Fields("categoria")), 1, 6)

            End If

            cajamarca_registro mytablex, mytabley, mytableb, mytablec
            mytablex.Update

        End If

        If Not mytablex.NoMatch Then
            mytablex.Edit
            mytablef.Seek "=", Mid$(Trim("" & mytabley.Fields("codigo")), 1, 15)

            If Not mytablef.NoMatch Then
                mytablex.Fields("familia") = Mid$(Trim("" & mytablef.Fields("categoria")), 1, 6)

            End If

            cajamarca_registro mytablex, mytabley, mytableb, mytablec
            mytablex.Update

        End If

        '---precios-----
        For j = 1 To 3
            mytablef.Seek "=", Mid$(Trim("" & mytabley.Fields("codigo")), 1, 15)

            If Not mytablef.NoMatch Then
                I = 0
                Do

                    If mytablef.EOF Then Exit Do
                    If Trim("" & mytablef.Fields("cod_articulo")) = Mid$(Trim("" & mytabley.Fields("codigo")), 1, 15) Then
                        '------------------------------
                        mytablea.Seek "=", Trim("" & mytablef.Fields("cod_articulo")), Format(j, "00")

                        If mytablea.NoMatch Then
                            mytablea.AddNew
                            I = I + 1
                            cajamarca_detalle01 mytablea, mytablef, I, j
                            mytablea.Update
         
                        End If

                        If Not mytablea.NoMatch Then
                            mytablea.Edit
                            I = I + 1
                            cajamarca_detalle01 mytablea, mytablef, I, j
                            mytablea.Update

                        End If

                        '------------------------------
                        Else: Exit Do

                    End If

                    mytablef.MoveNext
                Loop

            End If

        Next j

        mytabley.MoveNext
    Loop
    mytablex.Close
    mytabley.Close
    'mytablez.Close
    mytablea.Close
    mytableb.Close
    mytabled.Close
    'mytablec.Close
    mydby.Close

    MsgBox "Precios"
    ind = 1
    Set mytablea = mydbxglo.OpenTable("precios")
    mytablea.Index = "tprecios"
    Do

        If mytablea.EOF Then Exit Do
        vr = DoEvents()
        qap = "" & ind

        If Val("" & mytablea.Fields("factor1")) > 1 Then
            If Val("" & mytablea.Fields("factor2")) = 1 Then
                ind = ind + 1

                u1 = "" & mytablea.Fields("unidad1")
                f1 = Val("" & mytablea.Fields("factor1"))
                P1 = Val("" & mytablea.Fields("pventa1"))
                mytablea.Edit
                mytablea.Fields("unidad1") = "" & mytablea.Fields("unidad2")
                mytablea.Fields("factor1") = Val("" & mytablea.Fields("factor2"))
                mytablea.Fields("pventa1") = Val("" & mytablea.Fields("pventa2"))
                mytablea.Fields("unidad2") = u1
                mytablea.Fields("factor2") = f1
                mytablea.Fields("pventa2") = P1
                mytablea.Update
                GoTo akj

            End If

            If Val("" & mytablea.Fields("factor3")) = 1 Then
                ind = ind + 1

                u1 = "" & mytablea.Fields("unidad1")
                f1 = Val("" & mytablea.Fields("factor1"))
                P1 = Val("" & mytablea.Fields("pventa1"))
                mytablea.Edit
                mytablea.Fields("unidad1") = "" & mytablea.Fields("unidad3")
                mytablea.Fields("factor1") = Val("" & mytablea.Fields("factor3"))
                mytablea.Fields("pventa1") = Val("" & mytablea.Fields("pventa3"))
                mytablea.Fields("unidad3") = u1
                mytablea.Fields("factor3") = f1
                mytablea.Fields("pventa3") = P1
                mytablea.Update
                GoTo akj

            End If

        End If

akj:
        mytablea.MoveNext
    Loop
    mytablea.Close
    MsgBox "proceso Terminado", 48, "Aviso"

End Sub

Sub cajamarca_registro(mytablex As Table, _
                       mytabley As Table, _
                       mytableb As Table, _
                       mytablec As Table)
    mytablex.Fields("producto") = Mid$(Trim("" & mytabley.Fields("codigo")), 1, 15)
    mytablex.Fields("barras") = Trim("" & mytabley.Fields("codigo_bar"))
    mytablex.Fields("descripcio") = Mid$(Trim("" & mytabley.Fields("nombre")), 1, 60)
    mytablex.Fields("descorto") = Trim(Mid$("" & mytabley.Fields("nombre"), 1, 20))
    mytablex.Fields("oferta") = "" '& mytabley.Fields("flagremate")
    'mytablex.Fields("familia") = Mid$("" & mytabley.Fields("categoria"), 1, 6)
    mytablex.Fields("subfamilia") = "" 'Mid$("" & mytabley.Fields("descripcio"), 1, 2)
    mytablex.Fields("seccion") = "" '& mytabley.Fields("seccion")
    mytablex.Fields("marca") = "" & Mid$("" & mytabley.Fields("marca"), 1, 6)
    mytablex.Fields("categoria") = "" '& mytabley.Fields("categoria")
    mytablex.Fields("linea") = "" '& mytabley.Fields("flagtalla")
    mytablex.Fields("color") = "" '& mytabley.Fields("color")
    mytablex.Fields("fabrica") = ""
    'mytablex.Fields("proveedor1") = "" & mytabley.Fields("proveedor")
    'mytablex.Fields("proveedor2") = ""
    'mytablex.Fields("codprov1") = "" & mytabley.Fields("cpa")
    'mytablex.Fields("codprov2") = ""
    mytablex.Fields("serie") = ""
    mytablex.Fields("peso") = "" '& mytabley.Fields("balanza")
    mytablex.Fields("servicio") = ""
    mytablex.Fields("vtaund") = "" '& mytabley.Fields("flagunidad")
    mytablex.Fields("vecaja") = "S"
    mytablex.Fields("igv") = 19 'Val("" & mytabley.Fields("igv"))
    mytablex.Fields("isc") = 0
    mytablex.Fields("pesokgr") = 0.001
    mytablex.Fields("comision") = 0
    mytablex.Fields("monedac") = "S" '& mytabley.Fields("monedac")
    mytablex.Fields("unidad") = "UND" '& mytabley.Fields("unidad")
    mytablex.Fields("factor") = 1 'Val("" & mytabley.Fields("factor"))
    mytablex.Fields("costou") = 0 'Val("" & mytabley.Fields("costopaqu"))
    mytablex.Fields("costop") = 0 'Val("" & mytabley.Fields("costopaqp"))
    mytablex.Fields("monedav") = "S" ' & mytabley.Fields("moneda")

    'grabando precios al local nro 1
    'grabando almacen

    'mytablec.Seek "=", "" & mytabley.Fields("producto"), "01"
    'If Not mytablec.NoMatch Then
    'mytableb.Seek "=", "01", "" & mytablec.Fields("producto"), almacen
    'If mytableb.NoMatch Then
    '   mytableb.AddNew
    '   mytableb.Fields("local") = "01"
    '   mytableb.Fields("producto") = "" & mytablec.Fields("producto")
    '   mytableb.Fields("bodega") = almacen
    '   mytableb.Fields("saldo") = Val("" & mytablec.Fields("saldo"))
    '   mytableb.Update
    'End If
    'If Not mytableb.NoMatch Then
    '   mytableb.Edit
    '   mytableb.Fields("local") = "01"
    '   mytableb.Fields("producto") = "" & mytablec.Fields("producto")
    '   mytableb.Fields("bodega") = almacen
    '   mytableb.Fields("saldo") = Val("" & mytablec.Fields("saldo"))
    '   mytableb.Update
    'End If
    'End If

End Sub

Sub cajamarca_detalle01(mytablea As Table, mytablef As Table, I As Integer, sw As Integer)

    Dim buf As String

    If sw = 1 Then
        buf = Val("" & mytablef.Fields("preci1_mn"))

    End If

    If sw = 2 Then
        buf = Val("" & mytablef.Fields("preci2_mn"))

    End If

    If sw = 3 Then
        buf = Val("" & mytablef.Fields("preci3_mn"))

    End If

    mytablea.Fields("local") = Format(sw, "00")
    mytablea.Fields("producto") = Trim("" & mytablef.Fields("cod_articulo"))

    Select Case I

        Case 1
            mytablea.Fields("factor1") = Val("" & mytablef.Fields("equivalent"))
            mytablea.Fields("unidad1") = Mid$("" & mytablef.Fields("presentaci"), 1, 6)
            mytablea.Fields("pventa1") = Val(buf)

        Case 2
            mytablea.Fields("factor2") = Val("" & mytablef.Fields("equivalent"))
            mytablea.Fields("unidad2") = Mid$("" & mytablef.Fields("presentaci"), 1, 6)
            mytablea.Fields("pventa2") = Val(buf)

        Case 3
            mytablea.Fields("factor3") = Val("" & mytablef.Fields("equivalent"))
            mytablea.Fields("unidad3") = Mid$("" & mytablef.Fields("presentaci"), 1, 6)
            mytablea.Fields("pventa3") = Val(buf)

        Case 4
            mytablea.Fields("factor4") = Val("" & mytablef.Fields("equivalent"))
            mytablea.Fields("unidad4") = Mid$("" & mytablef.Fields("presentaci"), 1, 6)
            mytablea.Fields("pventa4") = Val(buf)

    End Select

End Sub

Sub graba_dona()

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset  'productos

    Dim mytabley As Table

    Dim vr

    Dim sdx      As Double

    Dim mytableb As Table 'almacen

    Dim mytablec As Table 'almacen orion ant

    sdx = 0

    If procesar <> "Dona" Then
        procesar = ""
        procesar.SetFocus
        Exit Sub

    End If

    'MsgBox orionv4
    'cn.Execute ("delete from producto")
    'cn.Execute ("delete from precios")
    'cn.Execute ("delete from dueno")
    'cn.Execute ("delete from codprov")

    orionv4 = "\precios"
    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("FINAL")

    Do

        If mytabley.EOF Then Exit Do
        mytablex.Open "select * from producto where producto='" & Trim("" & mytabley.Fields("producto")) & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            '   mytablex.AddNew
            '   pone_registro_dona mytablex, mytabley
            '   mytablex.Update
        Else
            pone_registro_dona1 mytablex, mytabley

            'mytablex.Update
        End If

        mytablex.Close
        sdx = sdx + 1
        vr = DoEvents()
        dd = "" & sdx
        mytabley.MoveNext
    Loop
    mydby.Close
    MsgBox "proceso Terminado", 48, "Aviso"

End Sub

Sub pone_registro_dona(mytablex As ADODB.Recordset, mytabley As Table)

    Dim mytablea As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    mytablex.Fields("producto") = Trim("" & mytabley.Fields("producto"))
    mytablex.Fields("barras") = Trim("" & mytabley.Fields("barras"))
    mytablex.Fields("descripcio") = Trim("" & mytabley.Fields("descripcio"))
    mytablex.Fields("descorto") = Mid$(Trim("" & mytabley.Fields("descripcio")), 1, 20)
    mytablex.Fields("presenta") = Trim("" & mytabley.Fields("presentaci"))
    mytablex.Fields("familia") = Trim("" & mytabley.Fields("famili"))
    mytablex.Fields("subfamilia") = Trim("" & mytabley.Fields("subfam"))
    mytablex.Fields("marca") = "" & mytabley.Fields("marca")
    mytablex.Fields("fabrica") = ""
    mytablex.Fields("serie") = ""
    mytablex.Fields("peso") = "N"
    mytablex.Fields("servicio") = ""
    mytablex.Fields("vecaja") = "S"
    mytablex.Fields("igv") = Val("" & mytabley.Fields("igv"))
    mytablex.Fields("isc") = 0
    mytablex.Fields("pesokgr") = 0.001
    mytablex.Fields("comision") = 0
    mytablex.Fields("monedac") = "S"
    mytablex.Fields("unidad") = "UND" '& mytabley.Fields("unidad")
    mytablex.Fields("factor") = 1
    mytablex.Fields("costou") = Val("" & mytabley.Fields("costo"))
    mytablex.Fields("costop") = Val("" & mytabley.Fields("costo"))
    mytablex.Fields("monedav") = "S"

    'grabando precios al local nro 1
    mytablea.Open "select * from precios where producto='" & Trim("" & mytabley.Fields("producto")) & "' and local='01'", cn, adOpenStatic, adLockOptimistic

    If mytablea.RecordCount = 0 Then
        mytablea.AddNew
        pone_detalle001 mytablea, mytabley, "01"
        mytablea.Update
    Else

        'pone_detalle001 mytablea, mytabley, "01"
        'mytablea.Update
    End If

    mytablea.Close

    mytablea.Open "select * from precios where producto='" & Trim("" & mytabley.Fields("producto")) & "' and local='02'", cn, adOpenStatic, adLockOptimistic

    If mytablea.RecordCount = 0 Then
        mytablea.AddNew
        pone_detalle001 mytablea, mytabley, "02"
        mytablea.Update
    Else

        'pone_detalle001 mytablea, mytabley, "02"
        'mytablea.Update
    End If

    mytablea.Close

    mytablea.Open "select * from precios where producto='" & Trim("" & mytabley.Fields("producto")) & "' and local='03'", cn, adOpenStatic, adLockOptimistic

    If mytablea.RecordCount = 0 Then
        mytablea.AddNew
        pone_detalle001 mytablea, mytabley, "01"
        mytablea.Update
    Else

        'pone_detalle001 mytablea, mytabley, "03"
        'mytablea.Update
    End If

    mytablea.Close

    mytablea.Open "select * from precios where producto='" & Trim("" & mytabley.Fields("producto")) & "' and local='04'", cn, adOpenStatic, adLockOptimistic

    If mytablea.RecordCount = 0 Then
        mytablea.AddNew
        pone_detalle001 mytablea, mytabley, "04"
        mytablea.Update
    Else

        'pone_detalle001 mytablea, mytabley, "04"
        'mytablea.Update
    End If

    mytablea.Close

End Sub

Sub pone_registro_dona1(mytablex As ADODB.Recordset, mytabley As Table)

    Dim mytablea As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    'mytablex.Fields("descripcio") = Trim("" & mytabley.Fields("descripcio"))

    'grabando precios al local nro 1
    'mytablea.Open "select * from precios where producto='" & Trim("" & mytabley.Fields("producto")) & "' and local='01'", cn, adOpenStatic, adLockOptimistic
    'If mytablea.RecordCount = 0 Then
    '   mytablea.AddNew
    '   pone_detalle001 mytablea, mytabley, "01"
    '   mytablea.Update
    '   Else
    '   pone_detalle001 mytablea, mytabley, "01"
    '   mytablea.Update
    'End If
    'mytablea.Close
    'mytablea.Open "select * from precios where producto='" & Trim("" & mytabley.Fields("producto")) & "' and local='02'", cn, adOpenStatic, adLockOptimistic
    'If mytablea.RecordCount = 0 Then
    '   mytablea.AddNew
    '   pone_detalle001 mytablea, mytabley, "02"
    '   mytablea.Update
    '   Else
    '   pone_detalle001 mytablea, mytabley, "02"
    '   mytablea.Update
    'End If
    'mytablea.Close

    mytablea.Open "select * from precios where producto='" & Trim("" & mytabley.Fields("producto")) & "' and local='03'", cn, adOpenStatic, adLockOptimistic

    If mytablea.RecordCount = 0 Then
        mytablea.AddNew
        pone_detalle001 mytablea, mytabley, "01"
        mytablea.Update
    Else
        pone_detalle001 mytablea, mytabley, "01"
        mytablea.Update

    End If

    mytablea.Close

    mytablea.Open "select * from precios where producto='" & Trim("" & mytabley.Fields("producto")) & "' and local='04'", cn, adOpenStatic, adLockOptimistic

    If mytablea.RecordCount = 0 Then
        mytablea.AddNew
        pone_detalle001 mytablea, mytabley, "04"
        mytablea.Update
    Else
        pone_detalle001 mytablea, mytabley, "04"
        mytablea.Update

    End If

    mytablea.Close

End Sub

Sub pone_detalle001(mytablex As ADODB.Recordset, mytabley As Table, buf As String)

    mytablex.Fields("local") = buf
    mytablex.Fields("producto") = "" & mytabley.Fields("producto")

    If buf = "01" Then

        mytablex.Fields("factor1") = 1
        mytablex.Fields("unidad1") = "UND"
        mytablex.Fields("pventa1") = Val("" & mytabley.Fields("pventa1"))

    End If

    If buf = "02" Then

        mytablex.Fields("factor1") = 1
        mytablex.Fields("unidad1") = "UND"
        mytablex.Fields("pventa1") = Val("" & mytabley.Fields("pventa1"))

    End If

    If buf = "03" Then

        mytablex.Fields("factor1") = 1
        mytablex.Fields("unidad1") = "UND"
        mytablex.Fields("pventa1") = Val("" & mytabley.Fields("pventa2"))

    End If

    If buf = "04" Then

        mytablex.Fields("factor1") = 1
        mytablex.Fields("unidad1") = "UND"
        mytablex.Fields("pventa1") = Val("" & mytabley.Fields("pventa3"))

    End If

End Sub

Sub clientes_v5()

    Dim I        As Integer

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    cn.Execute ("delete from clientes")
    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("clientes")

    Do

        If mytabley.EOF Then Exit Do
        mytablex.Open "select * from clientes where codigo='" & "" & mytabley.Fields("codigo") & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.AddNew
            'For i = 0 To mytabley.Fields.count - 10
            'mytablex.Fields(i) = mytabley.Fields(i)
            'Next i

            mytablex.Fields("codigo") = "" & mytabley.Fields("codigo")
            mytablex.Fields("nombre") = "" & mytabley.Fields("nombre")
            mytablex.Fields("direccion") = "" & mytabley.Fields("direccion")
            mytablex.Fields("moneda") = "S"
            mytablex.Fields("tipo") = "O"
            mytablex.Update

        End If

        mytablex.Close
        mytabley.MoveNext
    Loop
    '------------------------------------- ------------

    MsgBox "proceso Terminado", 48, "Aviso"

End Sub

Sub cuentac_v5()

    Dim I        As Integer

    Dim j        As Integer

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    On Error GoTo cmd8900_err

    j = 0
    cn.Execute ("delete from cuentac")
    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("cuentac")

    mytablex.Open "select * from cuentac", cn, adOpenStatic, adLockOptimistic
    Do

        If mytabley.EOF Then Exit Do
        j = j + 1
   
        mytablex.AddNew
   
        For I = 0 To mytabley.Fields.count - 2
            mytablex.Fields(I) = mytabley.Fields(I)
        Next I
   
        'mytablex.Fields("TIPO") = Trim("" & mytabley.Fields("tipo"))
        'mytablex.Fields("serie") = Trim("" & mytabley.Fields("serie"))
        'mytablex.Fields("numero") = Trim("" & mytabley.Fields("numero"))
        'mytablex.Fields("cuota") = Trim("" & mytabley.Fields("cuota"))
   
        'mytablex.Fields("tipoclie") = Trim("" & mytabley.Fields("tipoclie"))
        'mytablex.Fields("codigo") = Trim("" & mytabley.Fields("codigo"))
        'mytablex.Fields("nombre") = Trim("" & mytabley.Fields("nombre"))
   
        'mytablex.Fields("fecha") = CVDate("" & mytabley.Fields("fecha"))
        'mytablex.Fields("fechav") = CVDate("" & mytabley.Fields("fechav"))
   
        'mytablex.Fields("moneda") = Trim("" & mytabley.Fields("moneda"))
        'mytablex.Fields("total") = Val("" & mytabley.Fields("total"))
        'mytablex.Fields("abono") = Val("" & mytabley.Fields("abono"))
        'mytablex.Fields("saldo") = Val("" & mytabley.Fields("saldo"))
        'mytablex.Fields("interes") = Val("" & mytabley.Fields("interes"))
        'mytablex.Fields("estado") = Trim("" & mytabley.Fields("estado"))
        'mytablex.Fields("vendedor") = Trim("" & mytabley.Fields("vendedor"))
        'mytablex.Fields("dias") = Val("" & mytabley.Fields("dias"))
        'mytablex.Fields("local") = Trim("" & mytabley.Fields("local"))
        'mytablex.Fields("usuario") = Trim("" & mytabley.Fields("usuario"))
        'mytablex.Fields("caja") = Trim("" & mytabley.Fields("caja"))
        'mytablex.Fields("turno") = Trim("" & mytabley.Fields("turno"))
        'mytablex.Fields("fpago") = Trim("" & mytabley.Fields("fpago"))
        'mytablex.Fields("x") = Trim("" & mytabley.Fields("x"))
        'mytablex.Fields("anticipo") = Trim("" & mytabley.Fields("anticipo"))
        'mytablex.Fields("observa") = Trim("" & mytabley.Fields("observa"))
        'mytablex.Fields("acu") = Trim("" & mytabley.Fields("acu"))
        'mytablex.Fields("grupo") = Trim("" & mytabley.Fields("grupo"))
        mytablex.Fields("local") = "01"
        mytablex.Update

        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    MsgBox "proceso Terminado", 48, "Aviso"
    Exit Sub
cmd8900_err:
    MsgBox "Aviso " + j + " " & error$
    Exit Sub

End Sub

Sub parameca_v5()

    Dim I        As Integer

    Dim j        As Integer

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    On Error GoTo cmd8900_err

    j = 0
    cn.Execute ("delete from parameca")
    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("parameca")
    mytablex.Open "select * from parameca", cn, adOpenStatic, adLockOptimistic

    Do

        If mytabley.EOF Then Exit Do
        j = j + 1

        If Val("" & mytabley.Fields("caja")) > 0 Then
            mytablex.AddNew

            For I = 0 To mytabley.Fields.count - 30
                'If mytabley.Fields(i).Type = 10 Then
                '   If mytablex.Fields(i).Type = 202 Then
                mytablex.Fields(I) = mytabley.Fields(I)
                '   End If
                'End If
            Next I

            mytablex.Update

        End If

        'MsgBox mytabley.Fields.count

        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    MsgBox "proceso Terminado", 48, "Aviso"
    Exit Sub
cmd8900_err:
    MsgBox "Aviso " + j + " " & error$
    Exit Sub

End Sub

Sub cuentacd_v5()

    Dim I        As Integer

    Dim j        As Integer

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    On Error GoTo cmd8900_err

    j = 0
    cn.Execute ("delete from cuentacd")
    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("cuentacd")
    mytablex.Open "select * from cuentacd", cn, adOpenStatic, adLockOptimistic

    Do

        If mytabley.EOF Then Exit Do
        j = j + 1
   
        mytablex.AddNew

        For I = 0 To mytabley.Fields.count - 2
            mytablex.Fields(I) = mytabley.Fields(I)
        Next I
   
        'mytablex.Fields("codigo") = Trim("" & mytabley.Fields("codigo"))
        'mytablex.Fields("TIPO") = Trim("" & mytabley.Fields("tipo"))
        'mytablex.Fields("serie") = Trim("" & mytabley.Fields("serie"))
        'mytablex.Fields("numero") = Trim("" & mytabley.Fields("numero"))
        'mytablex.Fields("acu") = Trim("" & mytabley.Fields("acu"))
        'mytablex.Fields("tipo1") = Trim("" & mytabley.Fields("tipo1"))
        'mytablex.Fields("serie1") = Trim("" & mytabley.Fields("serie1"))
        'mytablex.Fields("numero1") = Trim("" & mytabley.Fields("numero1"))
        'mytablex.Fields("cuota") = Trim("" & mytabley.Fields("cuota"))
        'mytablex.Fields("moneda") = Trim("" & mytabley.Fields("moneda"))
        'mytablex.Fields("total") = Val("" & mytabley.Fields("total"))
        'mytablex.Fields("paga") = Val("" & mytabley.Fields("paga"))
        'mytablex.Fields("estado") = Trim("" & mytabley.Fields("estado"))
        'mytablex.Fields("paridad") = Val("" & mytabley.Fields("paridad"))
        'mytablex.Fields("fecha") = CVDate("" & mytabley.Fields("fecha"))
        'mytablex.Fields("hora") = Trim("" & mytabley.Fields("hora"))
        'mytablex.Fields("usuario") = Trim("" & mytabley.Fields("usuario"))
        'mytablex.Fields("local") = Trim("" & mytabley.Fields("local"))
        'mytablex.Fields("local1") = Trim("" & mytabley.Fields("local1"))
        'mytablex.Fields("tipoclie") = Trim("" & mytabley.Fields("tipoclie"))
        'mytablex.Fields("caja") = Trim("" & mytabley.Fields("caja"))
        'mytablex.Fields("turno") = Trim("" & mytabley.Fields("turno"))
        mytablex.Fields("local") = "01"
   
        mytablex.Update
        'MsgBox mytabley.Fields.count
   
        'mytablex.Fields("TIPO") = Trim("" & mytabley.Fields("tipo"))
        'mytablex.Fields("serie") = Trim("" & mytabley.Fields("serie"))
        'mytablex.Fields("numero") = Trim("" & mytabley.Fields("numero"))
        'mytablex.Fields("cuota") = Trim("" & mytabley.Fields("cuota"))

        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    MsgBox "proceso Terminado", 48, "Aviso"
    Exit Sub
cmd8900_err:
    MsgBox "Aviso " & j & " " & error$
    Exit Sub

End Sub

Sub recibos_v5()

    Dim I        As Integer

    Dim j        As Integer

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    Dim buf      As String

    On Error GoTo cmd8900_err

    j = 0

    cn.Execute ("delete from recibo")
    mytablex.Open "select * from recibo ", cn, adOpenStatic, adLockOptimistic

    buf = "select * from recibo where "
    buf = buf & " fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
    'MsgBox buf

    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytabley = mydby.CreateSnapshot(buf)
    MsgBox "procesar"
    Do

        If mytabley.EOF Then Exit Do
        j = j + 1
   
        mytablex.AddNew

        For I = 0 To mytabley.Fields.count - 1
            mytablex.Fields(I) = mytabley.Fields(I)
        Next I

        'mytablex.Fields("codigo") = Trim("" & mytabley.Fields("codigo"))
        'mytablex.Fields("TIPO") = Trim("" & mytabley.Fields("tipo"))
        'mytablex.Fields("serie") = Trim("" & mytabley.Fields("serie"))
        'mytablex.Fields("numero") = Trim("" & mytabley.Fields("numero"))
        'mytablex.Fields("acu") = Trim("" & mytabley.Fields("acu"))
        'mytablex.Fields("tipo1") = Trim("" & mytabley.Fields("tipo1"))
        'mytablex.Fields("serie1") = Trim("" & mytabley.Fields("serie1"))
        'mytablex.Fields("numero1") = Trim("" & mytabley.Fields("numero1"))
        'mytablex.Fields("cuota") = Trim("" & mytabley.Fields("cuota"))
        'mytablex.Fields("moneda") = Trim("" & mytabley.Fields("moneda"))
        'mytablex.Fields("total") = Val("" & mytabley.Fields("total"))
        'mytablex.Fields("paga") = Val("" & mytabley.Fields("paga"))
        'mytablex.Fields("estado") = Trim("" & mytabley.Fields("estado"))
        'mytablex.Fields("paridad") = Val("" & mytabley.Fields("paridad"))
        'mytablex.Fields("fecha") = CVDate("" & mytabley.Fields("fecha"))
        'mytablex.Fields("hora") = Trim("" & mytabley.Fields("hora"))
        'mytablex.Fields("usuario") = Trim("" & mytabley.Fields("usuario"))
        'mytablex.Fields("local") = Trim("" & mytabley.Fields("local"))
        'mytablex.Fields("local1") = Trim("" & mytabley.Fields("local1"))
        'mytablex.Fields("tipoclie") = Trim("" & mytabley.Fields("tipoclie"))
        'mytablex.Fields("caja") = Trim("" & mytabley.Fields("caja"))
        'mytablex.Fields("turno") = Trim("" & mytabley.Fields("turno"))
        mytablex.Fields("local") = "01"
        'Select Case "" & mytabley.Fields("acu")
        'Case "81", "82", "83", "84", "85"
        'mytablex.Fields("servicio") = "V"
        'Case "35", "36", "37", "38", "39"
        'mytablex.Fields("servicio") = "W"
        'End Select
        'mytablex.Fields("local1") = "03"
        mytablex.Update
        'MsgBox mytabley.Fields.count
   
        'mytablex.Fields("TIPO") = Trim("" & mytabley.Fields("tipo"))
        'mytablex.Fields("serie") = Trim("" & mytabley.Fields("serie"))
        'mytablex.Fields("numero") = Trim("" & mytabley.Fields("numero"))
        'mytablex.Fields("cuota") = Trim("" & mytabley.Fields("cuota"))

        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    MsgBox "proceso Terminado", 48, "Aviso"
    Exit Sub
cmd8900_err:
    MsgBox "Aviso " & j & " " & error$
    Exit Sub

End Sub

Sub vendedor_v5()

    Dim I        As Integer

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    cn.Execute ("delete from VENDEDOR where codigo<>'VICKY'")
    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("vendedor")

    Do

        If mytabley.EOF Then Exit Do
        mytablex.Open "select * from vendedor where codigo='" & "" & mytabley.Fields("codigo") & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.AddNew
            mytablex.Fields("codigo") = "" & mytabley.Fields("codigo")
            mytablex.Fields("nombre") = "" & mytabley.Fields("nombre")
            mytablex.Fields("direccion") = "" & mytabley.Fields("direccion")
            mytablex.Fields("moneda") = "S"
            mytablex.Fields("tipo") = "O"
            mytablex.Fields("veclave") = "S"
            mytablex.Fields("vevend") = "S"
   
            mytablex.Fields("v1") = "S" '& mytabley.Fields("v1")
            mytablex.Fields("v2") = "S" '& mytabley.Fields("v2")
            mytablex.Fields("v3") = "S" '& mytabley.Fields("v3")
            mytablex.Fields("v4") = "S" '& mytabley.Fields("v4")
   
            mytablex.Fields("v5") = "S" '& mytabley.Fields("v5")
            mytablex.Fields("v6") = "S" '& mytabley.Fields("v6")
            mytablex.Fields("v7") = "S" '& mytabley.Fields("v7")
            mytablex.Fields("v8") = "S" '& mytabley.Fields("v8")
   
            mytablex.Fields("v9") = "S" '& mytabley.Fields("v9")
            mytablex.Fields("v10") = "S" '& mytabley.Fields("v10")
            mytablex.Fields("v11") = "S" '& mytabley.Fields("v11")
            mytablex.Fields("v12") = "S" '& mytabley.Fields("v12")
   
            mytablex.Fields("clave") = "" & mytabley.Fields("clave")
            mytablex.Fields("rw1") = "W" '& mytabley.Fields("rw1")
            mytablex.Fields("rw2") = "W" '& mytabley.Fields("rw2")
            mytablex.Fields("rw3") = "W" '& mytabley.Fields("rw3")
            mytablex.Fields("rw4") = "W" '& mytabley.Fields("rw4")
   
            mytablex.Fields("rw5") = "W" '& mytabley.Fields("rw5")
            mytablex.Fields("rw6") = "W" '& mytabley.Fields("rw6")
            mytablex.Fields("rw7") = "W" '& mytabley.Fields("rw7")
            mytablex.Fields("rw8") = "W" '& mytabley.Fields("rw8")
   
            mytablex.Fields("rw9") = "W" '& mytabley.Fields("rw9")
            mytablex.Fields("rw10") = "W" '& mytabley.Fields("rw10")
            mytablex.Fields("rw11") = "W" '& mytabley.Fields("rw11")
            mytablex.Fields("rw12") = "W" '& mytabley.Fields("rw12")
            mytablex.Update
        Else
   
            mytablex.Fields("clave") = "" & mytabley.Fields("clave")
            mytablex.Fields("veclave") = "S"
            mytablex.Fields("vevend") = "S"
   
            mytablex.Fields("v1") = "S" '& mytabley.Fields("v1")
            mytablex.Fields("v2") = "S" '& mytabley.Fields("v2")
            mytablex.Fields("v3") = "S" '& mytabley.Fields("v3")
            mytablex.Fields("v4") = "S" '& mytabley.Fields("v4")
   
            mytablex.Fields("v5") = "S" '& mytabley.Fields("v5")
            mytablex.Fields("v6") = "S" '& mytabley.Fields("v6")
            mytablex.Fields("v7") = "S" '& mytabley.Fields("v7")
            mytablex.Fields("v8") = "S" '& mytabley.Fields("v8")
   
            mytablex.Fields("v9") = "S" '& mytabley.Fields("v9")
            mytablex.Fields("v10") = "S" '& mytabley.Fields("v10")
            mytablex.Fields("v11") = "S" '& mytabley.Fields("v11")
            mytablex.Fields("v12") = "S" '& mytabley.Fields("v12")
   
            mytablex.Fields("rw1") = "W" '& mytabley.Fields("rw1")
            mytablex.Fields("rw2") = "W" '& mytabley.Fields("rw2")
            mytablex.Fields("rw3") = "W" '& mytabley.Fields("rw3")
            mytablex.Fields("rw4") = "W" '& mytabley.Fields("rw4")
   
            mytablex.Fields("rw5") = "W" '& mytabley.Fields("rw5")
            mytablex.Fields("rw6") = "W" '& mytabley.Fields("rw6")
            mytablex.Fields("rw7") = "W" '& mytabley.Fields("rw7")
            mytablex.Fields("rw8") = "W" '& mytabley.Fields("rw8")
   
            mytablex.Fields("rw9") = "W" '& mytabley.Fields("rw9")
            mytablex.Fields("rw10") = "W" '& mytabley.Fields("rw10")
            mytablex.Fields("rw11") = "W" '& mytabley.Fields("rw11")
            mytablex.Fields("rw12") = "W" '& mytabley.Fields("rw12")
            mytablex.Update
   
        End If

        mytablex.Close
        mytabley.MoveNext
    Loop
    '------------------------------------- ------------

    MsgBox "proceso Terminado", 48, "Aviso"

End Sub

Sub tipodoc_v5()

    Dim I        As Integer

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    cn.Execute ("delete from tipo")
    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("tipo")
    Do

        If mytabley.EOF Then Exit Do
        mytablex.Open "select * from tipo where tipo='" & "" & mytabley.Fields("tipo") & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.AddNew
            'For i = 0 To mytabley.Fields.count - 10
            ' mytablex.Fields(i) = mytabley.Fields(i)
            mytablex.Fields("tipo") = "" & mytabley.Fields("tipo")
            mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcio")
            mytablex.Fields("tipodoc") = "" & mytabley.Fields("tipodoc")
            'Next i
            mytablex.Update

        End If

        mytablex.Close
        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    MsgBox "proceso Terminado", 48, "Aviso"

End Sub

Sub graba_v5producto()

    Dim mytablea As New ADODB.Recordset

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset  'productos

    Dim mytabley As Table

    Dim vr

    Dim sdx As Double

    cn.Execute ("delete from producto")
    'cn.Execute ("delete from precios")

    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("producto")
    mytablex.Open "select * from producto ", cn, adOpenStatic, adLockOptimistic

    Do

        If mytabley.EOF Then Exit Do
        'If mytablex.State = 1 Then
        '   mytablex.Close
        '   Set mytablex = Nothing
        'End If
        'mytablex.Open "select * from producto where producto='" & "" & mytabley.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic
        'If mytablex.RecordCount > 0 Then
        'pone_v5registro mytablex, mytabley
        'mytablex.Update
        'Else
        mytablex.AddNew
        pone_v5registro mytablex, mytabley
        mytablex.Update
        'End If
        'mytablex.Close
        sdx = sdx + 1
        vr = DoEvents()
        dd = "" & sdx
        mytabley.MoveNext
    Loop
    mytabley.Close
    mydby.Close
    MsgBox "proceso Terminado", 48, "Aviso"

End Sub

Sub pone_v5precios()

    Dim sdx      As Double

    Dim mydby    As Database

    Dim mytabley As Table

    Dim vr

    sdx = 0
    cn.Execute ("delete from precios")
    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("precios")
    Do

        If mytabley.EOF Then Exit Do
        pone_v5detalle01 mytabley, "01"
        sdx = sdx + 1
        vr = DoEvents()
        dd = "" & sdx
        mytabley.MoveNext
    Loop
    mytabley.Close
    mydby.Close

End Sub

Sub pone_v5registro(mytablex As ADODB.Recordset, mytabley As Table)
    mytablex.Fields("producto") = Trim("" & mytabley.Fields("producto"))
    mytablex.Fields("barras") = mytabley.Fields("barras")
    mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcio")
    mytablex.Fields("descorto") = "" & mytabley.Fields("descorto")
    mytablex.Fields("presenta") = "" & mytabley.Fields("presenta")
    mytablex.Fields("familia") = "" & mytabley.Fields("familia")
    mytablex.Fields("subfamilia") = "" & mytabley.Fields("subfamilia")
    mytablex.Fields("seccion") = "" & mytabley.Fields("seccion")
    mytablex.Fields("marca") = "" & mytabley.Fields("marca")
    mytablex.Fields("categoria") = "" & mytabley.Fields("categoria")
    mytablex.Fields("linea") = ""
    mytablex.Fields("color") = ""
    mytablex.Fields("fabrica") = ""
    mytablex.Fields("serie") = ""
    mytablex.Fields("peso") = ""
    mytablex.Fields("servicio") = ""
    mytablex.Fields("vecaja") = "S"
    mytablex.Fields("igv") = Val("" & mytabley.Fields("igv"))
    mytablex.Fields("isc") = Val("" & mytabley.Fields("isc"))
    mytablex.Fields("ivap") = Val("" & mytabley.Fields("ivap"))
    mytablex.Fields("pesokgr") = 0.001
    mytablex.Fields("comision") = 0
    mytablex.Fields("monedac") = "" & mytabley.Fields("monedac")
    mytablex.Fields("unidad") = "" & mytabley.Fields("unidad")
    mytablex.Fields("factor") = Val("" & mytabley.Fields("factor"))
    mytablex.Fields("costou") = Val("" & mytabley.Fields("costou"))
    mytablex.Fields("costop") = Val("" & mytabley.Fields("costop"))
    mytablex.Fields("monedav") = "" & mytabley.Fields("monedaV")
    mytablex.Fields("estado") = "S"
    mytablex.Fields("isc") = Val("" & mytabley.Fields("ISC"))
    mytablex.Fields("ivap") = Val("" & mytabley.Fields("IVAP")) 'ivap

    'grabando precios al local nro 1
End Sub

Sub pone_v5detalle01(mytabley As Table, buf As String)

    Dim mytablebb As New ADODB.Recordset

    Dim mytablez  As New ADODB.Recordset

    On Error GoTo cmd9012_err

    mytablebb.Open "select * from precios where producto='" & Trim("" & mytabley.Fields("producto")) & "' and local='01'", cn, adOpenStatic, adLockOptimistic

    If mytablebb.RecordCount = 0 Then
        mytablebb.AddNew
        mytablebb.Fields("local") = buf
        mytablebb.Fields("producto") = Trim("" & mytabley.Fields("producto"))
        mytablebb.Fields("ccosto") = "" & mytabley.Fields("ccosto")
        'mytablebb.Fields("monedav") = "" & mytabley.Fields("moneda")
        mytablebb.Fields("factor1") = Val("" & mytabley.Fields("factor1"))
        mytablebb.Fields("unidad1") = "" & mytabley.Fields("unidad1")
        mytablebb.Fields("pventa1") = Val("" & mytabley.Fields("pventa1"))

        mytablebb.Fields("factor2") = Val("" & mytabley.Fields("factor2"))
        mytablebb.Fields("unidad2") = "" & mytabley.Fields("unidad2")
        mytablebb.Fields("pventa2") = Val("" & mytabley.Fields("pventa2"))

        mytablebb.Fields("factor3") = Val("" & mytabley.Fields("factor3"))
        mytablebb.Fields("unidad3") = "" & mytabley.Fields("unidad3")
        mytablebb.Fields("pventa3") = Val("" & mytabley.Fields("pventa3"))

        mytablebb.Fields("factor4") = Val("" & mytabley.Fields("factor4"))
        mytablebb.Fields("unidad4") = "" & mytabley.Fields("unidad4")
        mytablebb.Fields("pventa4") = Val("" & mytabley.Fields("pventa4"))

        mytablebb.Fields("factor5") = Val("" & mytabley.Fields("factor5"))
        mytablebb.Fields("unidad5") = "" & mytabley.Fields("unidad5")
        mytablebb.Fields("pventa5") = Val("" & mytabley.Fields("pventa5"))

        mytablebb.Fields("factor6") = Val("" & mytabley.Fields("factor6"))
        mytablebb.Fields("unidad6") = "" & mytabley.Fields("unidad6")
        mytablebb.Fields("pventa6") = Val("" & mytabley.Fields("pventa6"))

        mytablebb.Fields("factor7") = Val("" & mytabley.Fields("factor7"))
        mytablebb.Fields("unidad7") = "" & mytabley.Fields("unidad7")
        mytablebb.Fields("pventa7") = Val("" & mytabley.Fields("pventa7"))

        mytablebb.Fields("factor8") = Val("" & mytabley.Fields("factor8"))
        mytablebb.Fields("unidad8") = "" & mytabley.Fields("unidad8")
        mytablebb.Fields("pventa8") = Val("" & mytabley.Fields("pventa8"))

        mytablebb.Fields("factor9") = Val("" & mytabley.Fields("factor9"))
        mytablebb.Fields("unidad9") = "" & mytabley.Fields("unidad9")
        mytablebb.Fields("pventa9") = Val("" & mytabley.Fields("pventa9"))

        mytablebb.Fields("factor10") = Val("" & mytabley.Fields("factor10"))
        mytablebb.Fields("unidad10") = "" & mytabley.Fields("unidad10")
        mytablebb.Fields("pventa10") = Val("" & mytabley.Fields("pventa10"))

        mytablebb.Fields("minimo11") = Val("" & mytabley.Fields("minimo11"))
        mytablebb.Fields("minimo12") = Val("" & mytabley.Fields("minimo12"))
        mytablebb.Fields("minimo13") = Val("" & mytabley.Fields("minimo13"))
        mytablebb.Fields("minimo14") = Val("" & mytabley.Fields("minimo14"))

        mytablebb.Fields("maximo11") = Val("" & mytabley.Fields("maximo11"))
        mytablebb.Fields("maximo12") = Val("" & mytabley.Fields("maximo12"))
        mytablebb.Fields("maximo13") = Val("" & mytabley.Fields("maximo13"))
        mytablebb.Fields("maximo14") = Val("" & mytabley.Fields("maximo14"))

        mytablebb.Fields("pventa11") = Val("" & mytabley.Fields("pventa11"))
        mytablebb.Fields("pventa12") = Val("" & mytabley.Fields("pventa12"))
        mytablebb.Fields("pventa13") = Val("" & mytabley.Fields("pventa13"))
        mytablebb.Fields("pventa14") = Val("" & mytabley.Fields("pventa14"))
        mytablebb.Update

    End If

    mytablebb.Close
    Exit Sub
cmd9012_err:
    MsgBox "" & mytabley.Fields("producto")
    Exit Sub

End Sub

Sub pone_v5familia()

    Dim sdx      As Double

    Dim mydby    As Database

    Dim mytabley As Table

    Dim mytablex As New ADODB.Recordset

    Dim vr

    sdx = 0
    cn.Execute ("delete from familia")
    mytablex.Open "select * from familia", cn, adOpenStatic, adLockOptimistic

    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("familia")
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew
        mytablex.Fields("familia") = Trim("" & mytabley.Fields("familia"))
        mytablex.Fields("descripcio") = Trim("" & mytabley.Fields("descripcio"))
        mytablex.Update
        sdx = sdx + 1
        vr = DoEvents()
        dd = "" & sdx
        mytabley.MoveNext
    Loop
    mytabley.Close
    mydby.Close

End Sub

Sub pone_v5subfamilia()

    Dim sdx      As Double

    Dim mydby    As Database

    Dim mytabley As Table

    Dim mytablex As New ADODB.Recordset

    Dim vr

    sdx = 0
    cn.Execute ("delete from subfamil")
    mytablex.Open "select * from subfamil", cn, adOpenStatic, adLockOptimistic

    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("subfamil")
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew
        mytablex.Fields("subfamilia") = Trim("" & mytabley.Fields("subfamilia"))
        mytablex.Fields("familia") = Trim("" & mytabley.Fields("familia"))
        mytablex.Fields("descripcio") = Trim("" & mytabley.Fields("descripcio"))
        mytablex.Update
        sdx = sdx + 1
        vr = DoEvents()
        dd = "" & sdx
        mytabley.MoveNext
    Loop
    mytabley.Close
    mydby.Close

End Sub

Sub pone_v5marca()

    Dim sdx      As Double

    Dim mydby    As Database

    Dim mytabley As Table

    Dim mytablex As New ADODB.Recordset

    Dim vr

    sdx = 0
    cn.Execute ("delete from marca")
    mytablex.Open "select * from marca", cn, adOpenStatic, adLockOptimistic

    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("marca")
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew
        mytablex.Fields("marca") = Trim("" & mytabley.Fields("marca"))
        mytablex.Fields("descripcio") = Trim("" & mytabley.Fields("descripcio"))
        mytablex.Update
        sdx = sdx + 1
        vr = DoEvents()
        dd = "" & sdx
        mytabley.MoveNext
    Loop
    mytabley.Close
    mydby.Close

End Sub

Sub pone_v5bodega()

    Dim sdx As Double

    Dim I   As Integer

    Dim vr

    Dim mydby    As Database

    Dim mytabley As Table

    Dim mytablex As New ADODB.Recordset

    sdx = 0
    cn.Execute ("delete from bodega")
    mytablex.Open "select * from bodega", cn, adOpenStatic, adLockOptimistic

    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("bodega")
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew

        For I = 0 To mytabley.Fields.count - 3
            mytablex.Fields(I) = mytabley.Fields(I)
        Next I

        mytablex.Update
        sdx = sdx + 1
        vr = DoEvents()
        dd = "" & sdx
        mytabley.MoveNext
    Loop
    mytabley.Close
    mydby.Close

End Sub

Sub pone_v5fpago()

    Dim sdx As Double

    Dim I   As Integer

    Dim vr

    Dim mydby    As Database

    Dim mytabley As Table

    Dim mytablex As New ADODB.Recordset

    sdx = 0
    cn.Execute ("delete from fpago")
    mytablex.Open "select * from fpago", cn, adOpenStatic, adLockOptimistic

    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("fpago")
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew

        For I = 0 To mytabley.Fields.count - 4
            mytablex.Fields(I) = mytabley.Fields(I)
        Next I

        mytablex.Update
        sdx = sdx + 1
        vr = DoEvents()
        dd = "" & sdx
        mytabley.MoveNext
    Loop
    mytabley.Close
    mydby.Close

End Sub

Sub pone_v5saldoini()

    Dim sdx As Double

    Dim I   As Integer

    Dim vr

    Dim mydby    As Database

    Dim mytabley As Table

    Dim mytablex As New ADODB.Recordset

    sdx = 0
    cn.Execute ("delete from saldoini")
    mytablex.Open "select * from saldoini", cn, adOpenStatic, adLockOptimistic

    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("saldoini")
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew

        For I = 0 To mytabley.Fields.count - 2
            mytablex.Fields(I) = mytabley.Fields(I)
        Next I

        mytablex.Fields("local") = "01"
        'mytablex.Fields("cantidad1") = Val("" & mytablex.Fields("cantidad"))
        mytablex.Update
        sdx = sdx + 1
        vr = DoEvents()
        dd = "" & sdx
        mytabley.MoveNext
    Loop
    mytabley.Close
    mydby.Close

End Sub

'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
Sub sixto()

    Dim mytablea As New ADODB.Recordset

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset  'productos

    Dim mytabley As Table

    Dim vr

    Dim sdx As Double

    cn.Execute ("delete from producto")
    'cn.Execute ("delete from precios")
    sdx = 1
    Set mydby = OpenDatabase("C:\sixto", False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("xx")
    mytablex.Open "select * from producto ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew
        sixto1 mytablex, mytabley, sdx
        mytablex.Update
        sdx = sdx + 1
        vr = DoEvents()
        dd = "" & sdx
        mytabley.MoveNext
    Loop
    mytabley.Close
    mydby.Close
    MsgBox "proceso Terminado", 48, "Aviso"

End Sub

Sub sixto1(mytablex As ADODB.Recordset, mytabley As Table, sdx1 As Double)
    mytablex.Fields("producto") = Trim(Mid$("" & mytabley.Fields("campo1"), 1, 15))
    'mytablex.Fields("producto") = "" & sdx1
    mytablex.Fields("barras") = ""
    mytablex.Fields("descripcio") = Mid$("" & mytabley.Fields("campo4"), 1, 60)
    mytablex.Fields("descorto") = Mid$("" & mytabley.Fields("campo4"), 1, 20)
    mytablex.Fields("presenta") = Mid$("" & mytabley.Fields("campo4"), 1, 20)
    mytablex.Fields("familia") = Mid$("" & mytabley.Fields("campo1"), 1, 2)
    mytablex.Fields("subfamilia") = Mid$("" & mytabley.Fields("campo1"), 3, 2)
    mytablex.Fields("seccion") = ""
    mytablex.Fields("marca") = Mid$("" & mytabley.Fields("campo3"), 1, 6)
    mytablex.Fields("categoria") = ""
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
    mytablex.Fields("unidad") = Mid$("" & mytabley.Fields("campo5"), 1, 6)
    mytablex.Fields("factor") = 1
    mytablex.Fields("costou") = 0
    mytablex.Fields("costop") = 0
    mytablex.Fields("monedav") = "S"

    'grabando precios al local nro 1
End Sub

Sub sixtoprecios()

    Dim mytablea As New ADODB.Recordset

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset  'productos

    Dim mytabley As Table

    Dim vr

    Dim sdx As Double

    cn.Execute ("delete from precios")
    'cn.Execute ("delete from precios")
    sdx = 1
    Set mydby = OpenDatabase("C:\sixto", False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("xxp")
    mytablex.Open "select * from precios ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew
        sixto1p mytablex, mytabley, sdx
        mytablex.Update
        sdx = sdx + 1
        vr = DoEvents()
        dd = "" & sdx
        mytabley.MoveNext
    Loop
    mytabley.Close
    mydby.Close
    MsgBox "proceso Terminado", 48, "Aviso"

End Sub

Sub sixto1p(mytablex As ADODB.Recordset, mytabley As Table, sdx1 As Double)
    mytablex.Fields("local") = Mid$(Trim("" & mytabley.Fields("campo1")), 1, 2)
    mytablex.Fields("producto") = Trim("" & mytabley.Fields("campo2"))
    mytablex.Fields("ccosto") = ""
    mytablex.Fields("factor1") = 1
    mytablex.Fields("unidad1") = Mid$(Trim("" & mytabley.Fields("campo6")), 1, 6)
    mytablex.Fields("pventa1") = Val("" & mytabley.Fields("campo8"))

End Sub

Sub graba_receta()

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    If procesar <> "Procesar" Then
        procesar = ""
        procesar.SetFocus
        Exit Sub

    End If

    cn.Execute ("delete from receta")
    Set mydby = OpenDatabase(orionv4, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("receta")

    mytablex.Open "select * from receta ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew
        mytablex.Fields("nro") = "1"
        mytablex.Fields("producto") = "" & mytabley.Fields("producto")
        mytablex.Fields("productoi") = "" & mytabley.Fields("insumo")
        mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcio")
        mytablex.Fields("unidad") = "" & mytabley.Fields("unidad")
        mytablex.Fields("factor") = Val("" & mytabley.Fields("factor"))
        mytablex.Fields("cantidad") = Val("" & mytabley.Fields("cantidad"))
        mytablex.Fields("local") = "01"
        mytablex.Fields("fecha") = Format(Now, "dd/mm/yyyy")
        mytablex.Update
        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    mytablex.Close
    MsgBox "graba receta Terminado", 48, "Aviso"

End Sub

Sub MIGRARV5SQL()

End Sub

Sub graba_facturav5()

    Dim I As Integer

    Dim vr

    Dim buf      As String

    Dim sdx      As Double

    Dim mydby    As Database

    Dim mytablex As Table

    Dim mytabley As New ADODB.Recordset

    cn.Execute ("delete from factura")

    mytabley.Open "select * from factura  ", cn, adOpenStatic, adLockOptimistic
    sdx = 0
    buf = "select * from factura where "
    buf = buf & " fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"

    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytablex = mydby.CreateSnapshot(buf)

    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + 1
        dd = "" & sdx
        vr = DoEvents()
        'avisamosfac mytablex, mytabley
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 5
            mytabley.Fields(I) = mytablex.Fields(I)
            'campo_campo mytabley, mytablex, i
        Next I

        mytabley.Fields("local") = "01"
        mytabley.Update
        mytablex.MoveNext
    Loop

End Sub

Sub campo_campo(mytabley As ADODB.Recordset, mytablex As Table, I As Integer)

    On Error GoTo cmd9090_err

    If mytablex.Fields(I).Type = 8 Then  'fecha
        If IsDate(mytablex.Fields(I)) Then
            mytabley.Fields(I) = mytablex.Fields(I)

        End If

    Else
        mytabley.Fields(I) = mytablex.Fields(I)

    End If

    Exit Sub
cmd9090_err:
    MsgBox mytablex.Fields(I)
    Exit Sub

End Sub

Sub avisamosfac(mytablex As Table, mytabley As ADODB.Recordset)

    'On Error GoTo cmd9078_err
    Dim I As Integer

    mytabley.AddNew

    For I = 0 To mytablex.Fields.count - 5

        'If mytablex.Fields(i).Type = 8 Then  'fecha
        '   If IsDate(mytablex.Fields(i)) Then
        '      mytabley.Fields(mytablex.Fields(i).Name) = mytablex.Fields(i)
        '   End If
        '   Else
        If Not IsNull(mytablex.Fields(I)) Then
            If mytabley.Fields(mytablex.Fields(I).Type) = mytablex.Fields(I).Type Then
                mytabley.Fields(mytablex.Fields(I).Name) = mytablex.Fields(I)

            End If

        End If
   
    Next I

    mytabley.Update
    Exit Sub

    'cmd9078_err:
    'MsgBox "Aviso en avisosfac " + error$, 48, "Aviso"
    'Exit Sub
End Sub

Sub graba_detallev5()

    Dim I As Integer

    Dim vr

    Dim buf      As String

    Dim sdx      As Double

    Dim mydby    As Database

    Dim mytablex As Table

    Dim mytabley As New ADODB.Recordset

    cn.Execute ("delete from detalle")
    mytabley.Open "select * from detalle ", cn, adOpenStatic, adLockOptimistic
    sdx = 0
    buf = "select * from detalle where "
    buf = buf & " fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"

    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytablex = mydby.CreateSnapshot(buf)
    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + 1
        dd = "" & sdx
        vr = DoEvents()
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 5
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Fields("local") = "01"
        mytabley.Update
        mytablex.MoveNext
    Loop

End Sub

Sub graba_fpagovv5()

    Dim I As Integer

    Dim vr

    Dim buf      As String

    Dim sdx      As Double

    Dim mydby    As Database

    Dim mytablex As Table

    Dim mytabley As New ADODB.Recordset

    cn.Execute ("delete from fpagov")
    mytabley.Open "select * from fpagov", cn, adOpenStatic, adLockOptimistic
    sdx = 0
    buf = "select * from fpagov where "
    buf = buf & " fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"

    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytablex = mydby.CreateSnapshot(buf)
    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + 1
        dd = "" & sdx
        vr = DoEvents()
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 5
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Fields("local") = "01"

        Select Case "" & mytablex.Fields("tipo")

            Case "81", "82", "83", "84", "85"
                mytabley.Fields("servicio") = "V"

            Case "35", "36", "37", "38", "39"
                mytabley.Fields("servicio") = "W"

        End Select

        mytabley.Update
        mytablex.MoveNext
    Loop

End Sub

Sub graba_almacenv4()

    Dim I As Integer

    Dim vr

    Dim dd       As String

    Dim sdx      As Double

    Dim mydby    As Database

    Dim mytablex As Table

    Dim mytabley As New ADODB.Recordset

    cn.Execute ("delete from saldoini")
    sdx = 0
    Set mydby = OpenDatabase("\orion.v4\001d\01", False, False, "foxpro 2.5;")
    Set mytablex = mydby.OpenTable("almacen")

    mytabley.Open "select *  from saldoini  ", cn, adOpenStatic, adLockOptimistic
    MsgBox "Empezar..enter" + "" & mytablex.RecordCount
    Do

        If mytablex.EOF Then Exit Do
        '------------------------
        mytabley.AddNew
        mytabley.Fields("local") = "01"
        mytabley.Fields("familia") = "" 'Mid$("" & mytabley.Fields("familia"), 1, 6)
        mytabley.Fields("producto") = "" & mytablex.Fields("producto")
        mytabley.Fields("DESCRIPCIO") = ""
        mytabley.Fields("unidad") = "UND"
        mytabley.Fields("factor") = 1
        mytabley.Fields("bodega") = "01"
        mytabley.Fields("fecha") = Format(Now, "dd/mm/yyyy")
        mytabley.Fields("cantidad") = Val("" & mytablex.Fields("saldo"))
        mytabley.Update

        sdx = sdx + 1
        dd = "" & sdx
        vr = DoEvents()

        '------------------------
        mytablex.MoveNext
    Loop
    MsgBox "acabe"

End Sub

Sub monica_cuenta()

    Dim I As Integer

    Dim vr

    Dim dd       As String

    Dim sdx      As Double

    Dim mydby    As Database

    Dim mytablex As Table

    Dim mytabley As New ADODB.Recordset

    cn.Execute ("delete from cuentas")
    sdx = 0
    Set mydby = OpenDatabase("C:\monica85", False, False, "foxpro 2.5;")
    Set mytablex = mydby.OpenTable("cuentas")

    mytabley.Open "select *  from cuentas  ", cn, adOpenStatic, adLockOptimistic
    MsgBox "Empezar..enter" + "" & mytablex.RecordCount
    Do

        If mytablex.EOF Then Exit Do
        '------------------------
        mytabley.AddNew
        mytabley.Fields("codcta") = Trim("" & mytablex.Fields("cod_cta"))
        mytabley.Fields("descripcio") = Trim("" & mytablex.Fields("name_cta"))
        mytabley.Fields("orden_blce") = Val(Trim("" & mytablex.Fields("orden_blce")))
        mytabley.Fields("modifica") = Trim("" & mytablex.Fields("modifica"))
        mytabley.Fields("tipo_cta") = Trim("" & mytablex.Fields("tipo_cta"))
        mytabley.Fields("blce_total") = Val("" & mytablex.Fields("blce_total"))
        mytabley.Fields("fe_ult_tr") = Val("" & mytablex.Fields("fe_ult_tr"))
        mytabley.Fields("comen1") = Trim("" & mytablex.Fields("comen1"))
        mytabley.Fields("comen2") = Trim("" & mytablex.Fields("comen2"))
        mytabley.Fields("nivel_cta") = Trim("" & mytablex.Fields("nivel_cta"))
        'mytabley.Fields("flag_ruc") = Trim("" & mytablex.Fields("nit"))
        'mytabley.Fields("ruc") = Trim("" & mytablex.Fields("nro_nit"))

        sdx = sdx + 1
        dd = "" & sdx
        vr = DoEvents()

        '------------------------
        mytablex.MoveNext
    Loop
    MsgBox "acabe"

End Sub

Sub siscont_cuenta()

    Dim I As Integer

    Dim vr

    Dim dd       As String

    Dim sdx      As Double

    Dim mydby    As Database

    Dim mytablex As Table

    Dim mytabley As New ADODB.Recordset

    cn.Execute ("delete from cuentas")
    sdx = 0
    Set mydby = OpenDatabase("C:\siscontoro", False, False, "foxpro 2.5;")
    Set mytablex = mydby.OpenTable("mdh_plan")

    mytabley.Open "select *  from cuentas  ", cn, adOpenStatic, adLockOptimistic
    MsgBox "Empezar..enter" + "" & mytablex.RecordCount
    Do

        If mytablex.EOF Then Exit Do

        '------------------------
        Select Case Mid$(Trim("" & mytablex.Fields("cuenta")), 1, 1)

            Case "1", "2", "3", "4", "5", "6", "7"
                mytabley.AddNew
                mytabley.Fields("codcta") = Trim("" & mytablex.Fields("cuenta"))
                mytabley.Fields("descripcio") = Trim("" & mytablex.Fields("nombre"))
                mytabley.Fields("orden_blce") = 0
                mytabley.Fields("modifica") = ""

                If Trim("" & mytablex.Fields("bd")) = "1" Or Trim("" & mytablex.Fields("bd")) = "2" Then
                    mytabley.Fields("nivel_cta") = "S"
                Else
                    mytabley.Fields("nivel_cta") = "D"

                End If
   
                If Val(Mid$(Trim("" & mytablex.Fields("cuenta")), 1, 1)) <= 3 Then
                    mytabley.Fields("tipo_cta") = "A"

                End If

                If Val(Mid$(Trim("" & mytablex.Fields("cuenta")), 1, 1)) = 4 Then
                    mytabley.Fields("tipo_cta") = "P"

                End If

                If Val(Mid$(Trim("" & mytablex.Fields("cuenta")), 1, 1)) = 5 Then
                    mytabley.Fields("tipo_cta") = "C"

                End If

                If Val(Mid$(Trim("" & mytablex.Fields("cuenta")), 1, 1)) = 6 Then
                    mytabley.Fields("tipo_cta") = "G"

                End If

                If Val(Mid$(Trim("" & mytablex.Fields("cuenta")), 1, 1)) = 7 Then
                    mytabley.Fields("tipo_cta") = "V"

                End If
     
                'mytabley.Fields("tipo_cta") = Trim("" & mytablex.Fields("tipo_cta"))
                'mytabley.Fields("blce_total") = Val("" & mytablex.Fields("blce_total"))
                'mytabley.Fields("fe_ult_tr") = Val("" & mytablex.Fields("fe_ult_tr"))
                'mytabley.Fields("comen1") = Trim("" & mytablex.Fields("comen1"))
                'mytabley.Fields("comen2") = Trim("" & mytablex.Fields("comen2"))
                'mytabley.Fields("nivel_cta") = Trim("" & mytablex.Fields("nivel_cta"))
                'mytabley.Fields("flag_ruc") = Trim("" & mytablex.Fields("nit"))
                'mytabley.Fields("ruc") = Trim("" & mytablex.Fields("nro_nit"))
                mytabley.Update

        End Select

        sdx = sdx + 1
        dd = "" & sdx
        vr = DoEvents()

        '------------------------
        mytablex.MoveNext
    Loop
    MsgBox "acabe"

End Sub

Sub carga_paramv5()

    Dim sdx As Double

    Dim I   As Integer

    Dim vr

    Dim mydby    As Database

    Dim mytabley As Table

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd88122_err

    sdx = 0
    cn.Execute ("delete from parameca")
    mytablex.Open "select * from parameca", cn, adOpenStatic, adLockOptimistic
    'MsgBox orionv5
    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("parameca")
    'MsgBox "" & mytabley.RecordCount
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew

        For I = 0 To mytabley.Fields.count - 10
            mytablex.Fields(I) = mytabley.Fields(I)
        Next I

        mytablex.Update
        sdx = sdx + 1
        vr = DoEvents()
        dd = "" & sdx
        mytabley.MoveNext
    Loop
    mytabley.Close
    mydby.Close
    'MsgBox "Acabe.." & sdx
    Exit Sub
cmd88122_err:
    MsgBox error$, 48, "Aviso"
    Exit Sub

End Sub

Sub graba_denisse()

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim mytablea As New ADODB.Recordset

    Dim mytableb As New ADODB.Recordset

    Dim sdx      As Double

    Dim vr

    cn.Execute ("delete from producto")
    cn.Execute ("delete from productb")
    cn.Execute ("delete from precios")
    sdx = 1
    mytablea.Open "select * from precios ", cn, adOpenStatic, adLockOptimistic
    mytableb.Open "select * from productb ", cn, adOpenStatic, adLockOptimistic
    mytablex.Open "select * from producto ", cn, adOpenStatic, adLockOptimistic
    mytabley.Open "select * from articulos ", cn, adOpenStatic, adLockOptimistic

    pasar_familias

    Do

        If mytabley.EOF Then Exit Do
        'If mytablex.RecordCount = 0 Then
        mytablex.AddNew
        pone_denisse mytablex, mytabley, sdx
        mytablex.Update
   
        mytablea.AddNew
        pone_preciodenisse mytablea, mytabley, "01", sdx
        mytablea.Update
   
        pone_denisse_barras mytableb, mytabley, sdx

        'End If
        vr = DoEvents()
        fechai = "" & sdx
        '-------------------------
        sdx = sdx + 1
        mytabley.MoveNext
    Loop

End Sub

Sub pone_denisse(mytablex As ADODB.Recordset, mytabley As ADODB.Recordset, sdx As Double)
    mytablex.Fields("producto") = "" & sdx
    mytablex.Fields("barras") = Mid$(Trim("" & mytabley.Fields("codart")), 1, 15)
    mytablex.Fields("barras") = ""
    mytablex.Fields("descripcio") = Mid$("" & mytabley.Fields("denart"), 1, 60)
    mytablex.Fields("descorto") = Mid$("" & mytabley.Fields("dabart"), 1, 20)
    mytablex.Fields("presenta") = ""
    mytablex.Fields("dsctoref") = 0
    mytablex.Fields("familia") = pone_familias_denisse(mytabley) 'Mid$(Trim("" & mytabley.Fields("codart")), 1, 6)
    mytablex.Fields("subfamilia") = ""
    mytablex.Fields("seccion") = ""
    mytablex.Fields("marca") = ""
    mytablex.Fields("categoria") = ""
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
    mytablex.Fields("costou") = Val("" & mytabley.Fields("precpr"))
    mytablex.Fields("costop") = Val("" & mytabley.Fields("precpr"))
    mytablex.Fields("monedav") = "S"
    mytablex.Fields("estado") = "S"
    mytablex.Fields("minimo") = 10
    mytablex.Fields("maximo") = 20

End Sub

Sub pone_preciodenisse(mytablex As ADODB.Recordset, _
                       mytabley As ADODB.Recordset, _
                       buf As String, _
                       sdx As Double)
    mytablex.Fields("local") = buf
    mytablex.Fields("producto") = "" & sdx
    mytablex.Fields("ccosto") = "" '& mytabley.Fields("seccion")
    mytablex.Fields("factor1") = 1 'Val("" & mytabley.Fields("factor1"))
    mytablex.Fields("unidad1") = "UND"
    mytablex.Fields("pventa1") = Val("" & mytabley.Fields("prevta"))

End Sub

Sub pone_denisse_barras(mytableb As ADODB.Recordset, _
                        mytabley As ADODB.Recordset, _
                        sdx As Double)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from cod_barra where codart='" & "" & mytabley.Fields("codart") & "'", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytableb.AddNew
        mytableb.Fields("producto") = "" & sdx
        mytableb.Fields("barras") = Mid$(Trim("" & mytablex.Fields("codrel")), 6, 13)
        mytableb.Update
        mytablex.MoveNext
    Loop
    mytablex.Close

End Sub

Function pone_familias_denisse(mytabley As ADODB.Recordset) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from gruposff where codgru='" & Mid$("" & mytabley.Fields("codart"), 1, 2) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        pone_familias_denisse = Trim(Mid$(Trim("" & mytablex.Fields("dengru")), 1, 6))

    End If

    mytablex.Close

End Function

Sub pasar_familias()

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    cn.Execute ("delete from familia")
    mytabley.Open "select * from familia ", cn, adOpenStatic, adLockOptimistic
    mytablex.Open "select * from gruposff ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew
        mytabley.Fields("familia") = Trim(Mid$(Trim("" & mytablex.Fields("dengru")), 1, 6))
        mytabley.Fields("descripcio") = Trim(Mid$(Trim("" & mytablex.Fields("dengru")), 1, 15))
        mytabley.Update
        mytablex.MoveNext
    Loop
    mytablex.Close

End Sub

'esto es de orion que sus precios estan en 1
Sub proveedor_v5()

    Dim I        As Integer

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    cn.Execute ("delete from proveedo")
    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("proveedo")
    mytablex.Open "select * from proveedo ", cn, adOpenStatic, adLockOptimistic

    Do

        If mytabley.EOF Then Exit Do
        'mytablex.Open "select * from proveedo where codigo='" & "" & mytabley.Fields("codigo") & "'", cn, adOpenStatic, adLockOptimistic
        'If mytablex.RecordCount = 0 Then
        mytablex.AddNew
        'For i = 0 To mytabley.Fields.count - 10
        'mytablex.Fields(i) = mytabley.Fields(i)
        'Next i

        mytablex.Fields("codigo") = quitar_blanco("" & mytabley.Fields("codigo"))
        mytablex.Fields("nombre") = quitar_blanco("" & mytabley.Fields("nombre"))
        mytablex.Fields("direccion") = quitar_blanco("" & mytabley.Fields("direccion"))
        mytablex.Fields("moneda") = "S"
        mytablex.Fields("tipo") = "O"
        mytablex.Update
        'End If

        mytabley.MoveNext
    Loop
    '------------------------------------- ------------

    MsgBox "proceso Terminado", 48, "Aviso"

End Sub

Function quitar_blanco(buf As String) As String

    Dim I    As Integer

    Dim buf1 As String

    buf1 = ""

    If Len(Trim(buf)) = 0 Then
        quitar_blanco = buf
        Exit Function

    End If

    For I = 1 To Len(buf)

        If Mid$(buf, I, 1) = "'" Then
        Else
            buf1 = buf1 & Mid$(buf, I, 1)

        End If

    Next I

    quitar_blanco = buf1

End Function

Sub graba_equivav5()

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    cn.Execute ("delete from productb")
    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("productb")

    'mytablex.Index = "marca"

    mytablex.Open "select * from productb ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew
        mytablex.Fields("producto") = "" & mytabley.Fields("producto")
        mytablex.Fields("barras") = "" & mytabley.Fields("barras")
        'mytablex.Fields("local") = "" & mytabley.Fields("local")
        mytablex.Update
        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    mytablex.Close
    MsgBox "Equiva proceso Terminado", 48, "Aviso"

End Sub

Sub graba_almacenv5()

    Dim I As Integer

    Dim vr

    Dim sdx      As Double

    Dim mydby    As Database

    Dim mytablex As Table

    Dim mytabley As New ADODB.Recordset

    Dim buf      As String

    buf = "delete from saldoini "
    'buf = buf & "  fecha='" & Format(fechai, "YYYYMMDD") & "'"
    cn.Execute (buf)
    sdx = 0
    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytablex = mydby.OpenTable("almacen")
    dd.Enabled = True
    dd.Visible = True
    dd = ""
    mytabley.Open "select *  from saldoini where 1=2 ", cn, adOpenStatic, adLockOptimistic
    MsgBox "Empezar..enter" + "" & mytablex.RecordCount
    Do

        If mytablex.EOF Then Exit Do
        '------------------------
        sdx = sdx + 1
        vr = DoEvents()
        dd = "" & sdx
        mytabley.AddNew
        mytabley.Fields("local") = "01"
        mytabley.Fields("familia") = "" '& Mid$("" & mytablex.Fields("familia"), 1, 6)
        mytabley.Fields("producto") = "" & mytablex.Fields("producto")
        mytabley.Fields("DESCRIPCIO") = ""
        mytabley.Fields("unidad") = "UND"
        mytabley.Fields("factor") = 1
        mytabley.Fields("bodega") = "" & mytablex.Fields("bodega")
        mytabley.Fields("fecha") = Format(fechai, "dd/mm/yyyy")
        mytabley.Fields("cantidad") = Val("" & mytablex.Fields("saldo"))
        mytabley.Update

        '------------------------
        mytablex.MoveNext
    Loop
    MsgBox "acabe"

End Sub

'-------------------------------------------------------------------------
Sub graba_recave()

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset  'productos

    Dim mytabley As Table

    Dim vr

    Dim sdx As Double

    sdx = 0
    cn.Execute ("delete from familia")
    cn.Execute ("delete from producto")
    cn.Execute ("delete from precios")
    cn.Execute ("delete from dueno")
    cn.Execute ("delete from codprov")

    'MsgBox orionv4
    Set mydby = OpenDatabase("\RECAVE\MDB", False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("plu_jpgs.001")

    Do

        If mytabley.EOF Then Exit Do
        mytablex.Open "select * from producto where producto='" & "" & mytabley.Fields("plu_code") & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.AddNew
            pone_recave mytablex, mytabley
            mytablex.Update
        Else
            pone_recave mytablex, mytabley
            mytablex.Update

        End If

        mytablex.Close
        sdx = sdx + 1
        vr = DoEvents()
        dd = "" & sdx
        mytabley.MoveNext
    Loop
    mytabley.Close
    mydby.Close
    MsgBox "Producto proceso Terminado", 48, "Aviso"

End Sub

Sub pone_recave(mytablex As ADODB.Recordset, mytabley As Table)

    Dim mytablea As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    mytablex.Fields("isc") = 0
    mytablex.Fields("ivap") = 0

    mytablex.Fields("producto") = Mid$(Trim("" & mytabley.Fields("plu_code")), 1, 15)
    mytablex.Fields("barras") = ""
    mytablex.Fields("descripcio") = Mid$(Trim("" & mytabley.Fields("plu_name")), 1, 80)
    mytablex.Fields("descorto") = Mid$(Trim("" & mytabley.Fields("plu_ticket")), 1, 22)

    mytablex.Fields("presenta") = ""
    mytablex.Fields("dsctoref") = 0

    mytablex.Fields("familia") = Mid$(Trim("" & mytabley.Fields("div_name")), 1, 6)
    mytablex.Fields("subfamilia") = ""
    mytablex.Fields("seccion") = ""
    mytablex.Fields("marca") = ""
    mytablex.Fields("categoria") = ""
    'mytablex.Fields("linea") = "" & mytabley.Fields("flagtalla")
    mytablex.Fields("color") = ""
    mytablex.Fields("fabrica") = ""
    mytablex.Fields("serie") = ""
    mytablex.Fields("peso") = "N"
    mytablex.Fields("servicio") = ""
    mytablex.Fields("vecaja") = "S"
    mytablex.Fields("igv") = 19
    mytablex.Fields("isc") = 0
    mytablex.Fields("pesokgr") = 0.001
    mytablex.Fields("comision") = 0

    If Trim("" & mytabley.Fields("plu_moneda")) = "1" Then
        mytablex.Fields("monedac") = "S"
        mytablex.Fields("monedav") = "S"

    End If

    If Trim("" & mytabley.Fields("plu_moneda")) = "2" Then
        mytablex.Fields("monedac") = "D"
        mytablex.Fields("monedav") = "D"

    End If

    mytablex.Fields("unidad") = "UND"
    mytablex.Fields("factor") = 1
    mytablex.Fields("costou") = 0
    mytablex.Fields("costop") = 0
    mytablex.Fields("estado") = "S"

    mytablez.Open "select * from familia where familia='" & Mid$(Trim("" & mytabley.Fields("div_name")), 1, 6) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablez.RecordCount = 0 Then
        mytablez.AddNew
        mytablez.Fields("familia") = Mid$(Trim("" & mytabley.Fields("div_name")), 1, 6)
        mytablez.Fields("descripcio") = Mid$(Trim("" & mytabley.Fields("div_name")), 1, 30)
        mytablez.Update
    Else
        mytablez.Fields("familia") = Mid$(Trim("" & mytabley.Fields("div_name")), 1, 6)
        mytablez.Fields("descripcio") = Mid$(Trim("" & mytabley.Fields("div_name")), 1, 30)
        mytablez.Update

    End If

    mytablez.Close

End Sub

Sub graba_recave_equiva()

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    cn.Execute ("delete from productb")
    Set mydby = OpenDatabase("\RECAVE\MDB", False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("EAN_JPGS.001")
    mytablex.Open "select * from productb ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew
        mytablex.Fields("producto") = Mid$(Trim("" & mytabley.Fields("plu_code")), 1, 15)
        mytablex.Fields("barras") = Mid$(Trim("" & mytabley.Fields("ean_code")), 1, 15)
        mytablex.Update
        mytabley.MoveNext
    Loop
    mytablex.Close
    MsgBox "Equiva proceso Terminado", 48, "Aviso"

End Sub

Sub graba_recave_precios()

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    cn.Execute ("delete from precios")
    Set mydby = OpenDatabase("\RECAVE\MDB", False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("PRE_JPGS.001")
    mytablex.Open "select * from precios ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytabley.EOF Then Exit Do
        'If Trim("" & mytabley.Fields("are_code")) = "01" Then
        mytablex.AddNew
        mytablex.Fields("local") = "01"
        mytablex.Fields("producto") = Mid$(Trim("" & mytabley.Fields("plu_code")), 1, 15)
        mytablex.Fields("unidad1") = "UND"
        mytablex.Fields("factor1") = 1
        mytablex.Fields("pventa1") = Val("" & mytabley.Fields("plu_precio"))
        mytablex.Update
        'End If
        mytabley.MoveNext
    Loop
    mytablex.Close
    MsgBox "Precios proceso Terminado", 48, "Aviso"

End Sub

Sub graba_percepcion()

    Dim sdx As Double

    Dim vr

    Dim mytabley As New ADODB.Recordset

    Dim mytablex As New ADODB.Recordset

    mytabley.Open "select * from productoborrado ", cn, adOpenStatic, adLockOptimistic
    dd = ""
    sdx = 0
    Do

        If mytabley.EOF Then Exit Do
        mytablex.Open "select * from producto where producto='" & Trim("" & mytabley.Fields("producto")) & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            If Mid$(Trim("" & mytabley.Fields("percepcion")), 1, 1) = "S" Then
                mytablex.Fields("percepcion") = Mid$(Trim("" & mytabley.Fields("percepcion")), 1, 1)
                sdx = sdx + 1

            End If

            mytablex.Update

        End If

        mytablex.Close
        vr = DoEvents()
        dd = "" & sdx
        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    MsgBox "Percepcion proceso Terminado", 48, "Aviso"

End Sub

Sub graba_percepcion_clientes()

    Dim sdx As Double

    Dim vr

    Dim mytabley As New ADODB.Recordset

    Dim mytablex As New ADODB.Recordset

    mytabley.Open "select * from clientesborrado ", cn, adOpenStatic, adLockOptimistic
    dd = ""
    sdx = 0
    Do

        If mytabley.EOF Then Exit Do
        mytablex.Open "select * from clientes where codigo='" & Mid$(Trim("" & mytabley.Fields(0)), 1, 11) & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.AddNew
            mytablex.Fields("codigo") = Mid$(Trim("" & mytabley.Fields(0)), 1, 11)
            mytablex.Fields("nombre") = Mid$(Trim("" & mytabley.Fields(1)), 1, 60)
            mytablex.Fields("moneda") = "S"
            mytablex.Fields("tipo") = "J"
            mytablex.Fields("clasesunat") = Mid$(Trim("" & mytabley.Fields(2)), 1, 1)
            sdx = sdx + 1
            mytablex.Update

        End If

        If mytablex.RecordCount > 0 Then
            mytablex.Fields("codigo") = Mid$(Trim("" & mytabley.Fields(0)), 1, 11)
            mytablex.Fields("moneda") = "S"
            mytablex.Fields("tipo") = "J"
            mytablex.Fields("clasesunat") = Mid$(Trim("" & mytabley.Fields(2)), 1, 1)
            sdx = sdx + 1
            mytablex.Update

        End If

        mytablex.Close
        vr = DoEvents()
        dd = "" & sdx
        mytabley.MoveNext
    Loop
    '------------------------------------- ------------
    MsgBox "Clientes Percepcion proceso Terminado", 48, "Aviso"

End Sub

Sub graba_productolaritza()

    Dim mytablex As New ADODB.Recordset  'productos

    Dim mytabley As New ADODB.Recordset

    Dim vr

    Dim sdx As Double

    sdx = 0
    cn.Execute ("delete from producto")
    cn.Execute ("delete from PRECIOS")
    mytabley.Open "select * from depurada ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytabley.EOF Then Exit Do
        If IsNumeric(mytabley.Fields("F1")) Then
            mytablex.Open "select * from producto where producto='" & Trim("" & mytabley.Fields("F1")) & "'", cn, adOpenStatic, adLockOptimistic

            If mytablex.RecordCount = 0 Then
                mytablex.AddNew
                pone_registrolaritza mytablex, mytabley
                mytablex.Update
            Else
                pone_registrolaritza mytablex, mytabley
                mytablex.Update

            End If

            mytablex.Close

        End If

        sdx = sdx + 1
        vr = DoEvents()
        dd = "" & sdx
        mytabley.MoveNext
    Loop
    mytabley.Close

    MsgBox "Producto proceso Terminado", 48, "Aviso"

End Sub

Sub pone_registrolaritza(mytablex As ADODB.Recordset, mytabley As ADODB.Recordset)

    Dim mytablea As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    On Error GoTo CMD901212_ERR

    mytablex.Fields("isc") = 0 ' Val("" & mytabley.Fields("ISC"))
    'mytablex.Fields("ivap") = Val("" & mytabley.Fields("nodscto")) 'ivap
    'Exit Sub

    mytablex.Fields("producto") = Trim("" & mytabley.Fields("f1"))
    mytablex.Fields("barras") = "" '& mytabley.Fields("barras")
    mytablex.Fields("descripcio") = Mid$(Trim("" & mytabley.Fields("F4")), 1, 80)
    mytablex.Fields("descorto") = Mid$(Trim("" & mytabley.Fields("F4")), 1, 22)

    mytablex.Fields("presenta") = ""
    mytablex.Fields("dsctoref") = 0

    mytablex.Fields("familia") = Mid$(Trim("" & mytabley.Fields("F3")), 1, 6)
    mytablex.Fields("subfamilia") = ""
    mytablex.Fields("seccion") = ""
    mytablex.Fields("marca") = ""
    mytablex.Fields("categoria") = ""
    'mytablex.Fields("linea") = "" & mytabley.Fields("flagtalla")
    mytablex.Fields("color") = ""
    mytablex.Fields("fabrica") = ""
    mytablex.Fields("serie") = ""
    mytablex.Fields("peso") = "N"
    mytablex.Fields("servicio") = "N"
    mytablex.Fields("vecaja") = "S"
    mytablex.Fields("igv") = 18
    mytablex.Fields("isc") = 0
    mytablex.Fields("pesokgr") = 0.001
    mytablex.Fields("comision") = 0
    mytablex.Fields("monedac") = "S"
    mytablex.Fields("unidad") = "UND"
    mytablex.Fields("factor") = 1
    mytablex.Fields("costou") = 0#
    mytablex.Fields("costop") = 0#
    mytablex.Fields("monedav") = "S"
    mytablex.Fields("estado") = "S"
    'mytablex.Fields("minimo") = Val("" & mytabley.Fields("stkminimo"))
    'mytablex.Fields("maximo") = Val("" & mytabley.Fields("stkmaximo"))

    'mytablez.Open "select * from dueno where codigo='" & "" & mytabley.Fields("seccion") & "' and local='01' and producto='" & "" & mytabley.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic
    'If mytablez.RecordCount = 0 Then
    '   mytablez.AddNew
    '   mytablez.Fields("codigo") = "" & mytabley.Fields("seccion")
    '   mytablez.Fields("local") = "01"
    '   mytablez.Fields("producto") = "" & mytabley.Fields("producto")
    '   mytablez.Update
    'Else
   
    '   mytablez.Fields("codigo") = "" & mytabley.Fields("seccion")
    '   mytablez.Fields("local") = "01"
    '   mytablez.Fields("producto") = "" & mytabley.Fields("producto")
    '   mytablez.Update
    'End If
    'mytablez.Close
    'mytablez.Open "select * from codprov where codigo='" & "" & mytabley.Fields("proveedor") & "' and producto='" & mytabley.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic
    'If mytablez.RecordCount = 0 Then
    '   mytablez.AddNew
    '   pone_detalle mytablez, mytabley, 0
    '   mytablez.Update
    'Else
   
    '   pone_detalle mytablez, mytabley, 0
    '   mytablez.Update
    'End If
    'mytablez.Close

    'mytablez.Open "select * from codprov where codigo='" & "" & mytabley.Fields("proveedor1") & "' and producto='" & mytabley.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic
    'If mytablez.RecordCount = 0 Then
    '   mytablez.AddNew
    '   pone_detalle mytablez, mytabley, 1
    '   mytablez.Update
    'Else
 
    '   pone_detalle mytablez, mytabley, 1
    '   mytablez.Update
    'End If
    'mytablez.Close

    'grabando precios al local nro 1
    mytablea.Open "select * from precios where producto='" & Trim("" & mytabley.Fields("F1")) & "' and local='01'", cn, adOpenStatic, adLockOptimistic

    If mytablea.RecordCount = 0 Then
        mytablea.AddNew
        pone_detalle01laritza mytablea, mytabley, "01"
        mytablea.Update
    Else
        pone_detalle01laritza mytablea, mytabley, "01"
        mytablea.Update

    End If

    mytablea.Close
    'mytablea.Open "select * from precios where producto='" & "" & mytabley.Fields("producto") & "' and local='02'", cn, adOpenStatic, adLockOptimistic
    'If mytablea.RecordCount = 0 Then '
    '   mytablea.AddNew
    '   pone_detalle01 mytablea, mytabley, "02"
    '   mytablea.Update
    '   Else
    '   pone_detalle01 mytablea, mytabley, "02"
    '   mytablea.Update
    'End If
    'mytablea.Close
    'mytablea.Open "select * from precios where producto='" & "" & mytabley.Fields("producto") & "' and local='03'", cn, adOpenStatic, adLockOptimistic
    'If mytablea.RecordCount = 0 Then
    '   mytablea.AddNew
    '   pone_detalle01 mytablea, mytabley, "03"
    '   mytablea.Update
    '   Else
    '   pone_detalle01 mytablea, mytabley, "03"
    '   mytablea.Update
    'End If
    'mytablea.Close
    'mytablea.Open "select * from precios where producto='" & "" & mytabley.Fields("producto") & "' and local='04'", cn, adOpenStatic, adLockOptimistic
    'If mytablea.RecordCount = 0 Then
    '   mytablea.AddNew
    '   pone_detalle01 mytablea, mytabley, "04"
    '   mytablea.Update
    '   Else
    '   pone_detalle01 mytablea, mytabley, "04"
    '   mytablea.Update
    'End If
    'mytablea.Close

    'grabando almacen
    'mytablec.Seek "=", "" & mytabley.Fields("producto"), "01"
    'If Not mytablec.NoMatch Then
    'mytableb.Seek "=", "01", "" & mytablec.Fields("producto"), almacen
    'If mytableb.NoMatch Then
    '   mytableb.AddNew
    '   mytableb.Fields("local") = "01"
    '   mytableb.Fields("producto") = "" & mytablec.Fields("producto")
    '   mytableb.Fields("bodega") = almacen
    '   mytableb.Fields("saldo") = Val("" & mytablec.Fields("saldo"))
    '   mytableb.Update
    'End If
    'If Not mytableb.NoMatch Then
    '   mytableb.Edit
    '   mytableb.Fields("local") = "01"
    '   mytableb.Fields("producto") = "" & mytablec.Fields("producto")
    '   mytableb.Fields("bodega") = almacen
    '   mytableb.Fields("saldo") = Val("" & mytablec.Fields("saldo"))
    '   mytableb.Update
    'End If
    'End If
    Exit Sub
CMD901212_ERR:
    MsgBox "Aviso en poner registro laritza " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub pone_detalle01laritza(mytablex As ADODB.Recordset, _
                          mytabley As ADODB.Recordset, _
                          buf As String)
    mytablex.Fields("local") = buf
    mytablex.Fields("producto") = Trim("" & mytabley.Fields("f1"))
    mytablex.Fields("ccosto") = ""
    'mytablex.Fields("monedav") = "" & mytabley.Fields("moneda")
    mytablex.Fields("factor1") = 1
    mytablex.Fields("unidad1") = "UND"
    mytablex.Fields("pventa1") = Format(Val("" & mytabley.Fields("f5")), "000.00")

End Sub

Sub graba_familialaritza()

    Dim mydby    As Database

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    cn.Execute ("delete from familia")
    mytabley.Open "select * from DEPURADA ", cn, adOpenStatic, adLockOptimistic

    Do

        If mytabley.EOF Then Exit Do
        mytablex.Open "select * from familia where familia='" & Mid$(Trim("" & mytabley.Fields("f3")), 1, 6) & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.AddNew
            mytablex.Fields("familia") = Mid$(Trim("" & mytabley.Fields("f3")), 1, 6)
            mytablex.Fields("descripcio") = Mid$(Trim("" & mytabley.Fields("f3")), 1, 6)
            mytablex.Fields("vetouch") = "S"
            mytablex.Update
        Else
            mytablex.Fields("familia") = Mid$(Trim("" & mytabley.Fields("f3")), 1, 6)
            mytablex.Fields("descripcio") = Mid$(Trim("" & mytabley.Fields("f3")), 1, 6)
            mytablex.Fields("vetouch") = "S"
            mytablex.Update

        End If

        mytablex.Close
        mytabley.MoveNext
    Loop
    MsgBox "Familia laritza terminado"

    '------------------------------------- ------------
End Sub

Sub precios_asia()

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    cn.Execute ("delete from precios where local='05'")
    mytablex.Open "select * from hoja1", cn, adOpenStatic, adLockOptimistic
    mytabley.Open "select * from precios ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew
        mytabley.Fields("local") = "05"
        mytabley.Fields("producto") = Trim("" & mytablex.Fields("F0"))
        mytabley.Fields("ccosto") = ""
        mytabley.Fields("factor1") = 1
        mytabley.Fields("unidad1") = "UND"
        mytabley.Fields("pventa1") = Format(Val("" & mytablex.Fields("f3")), "000.00")
        mytabley.Update
        mytablex.MoveNext
    Loop
    MsgBox "abc"

End Sub

Sub graba_cproformv5()

    Dim I As Integer

    Dim vr

    Dim buf      As String

    Dim sdx      As Double

    Dim mydby    As Database

    Dim mytablex As Table

    Dim mytabley As New ADODB.Recordset

    cn.Execute ("delete from cproform")

    mytabley.Open "select * from cproform  ", cn, adOpenStatic, adLockOptimistic
    sdx = 0
    buf = "select * from cproform where "
    buf = buf & " fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"

    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytablex = mydby.CreateSnapshot(buf)

    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + 1
        dd = "" & sdx
        vr = DoEvents()
        'avisamosfac mytablex, mytabley
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 5
            mytabley.Fields(I) = mytablex.Fields(I)
            'campo_campo mytabley, mytablex, i
        Next I

        mytabley.Fields("local") = "01"
        mytabley.Update
        mytablex.MoveNext
    Loop

End Sub

Sub graba_dproformv5()

    Dim I As Integer

    Dim vr

    Dim buf      As String

    Dim sdx      As Double

    Dim mydby    As Database

    Dim mytablex As Table

    Dim mytabley As New ADODB.Recordset

    cn.Execute ("delete from dproform")
    mytabley.Open "select * from dproform ", cn, adOpenStatic, adLockOptimistic
    sdx = 0
    buf = "select * from dproform where "
    buf = buf & " fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"

    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytablex = mydby.CreateSnapshot(buf)
    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + 1
        dd = "" & sdx
        vr = DoEvents()
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 5
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Fields("local") = "01"
        mytabley.Update
        mytablex.MoveNext
    Loop

End Sub

Sub pone_v5almacen()

    Dim sdx As Double

    Dim I   As Integer

    Dim vr

    Dim mydby    As Database

    Dim mytabley As Table

    Dim mytablex As New ADODB.Recordset

    sdx = 0
    cn.Execute ("delete from almacen")
    mytablex.Open "select * from almacen", cn, adOpenStatic, adLockOptimistic

    Set mydby = OpenDatabase(orionv5, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("almacen")
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew

        For I = 0 To mytabley.Fields.count - 3
            mytablex.Fields(I) = mytabley.Fields(I)
        Next I

        mytablex.Update
        sdx = sdx + 1
        vr = DoEvents()
        dd = "" & sdx
        mytabley.MoveNext
    Loop
    mytabley.Close
    mydby.Close

End Sub

'---- David
Sub graba_david()

    Dim mytablex As New ADODB.Recordset  'productos

    Dim mytabley As New ADODB.Recordset

    Dim vr

    Dim sdx As Double

    sdx = 0
    cn.Execute ("delete from producto")
    cn.Execute ("delete from precios")
    cn.Execute ("delete from FAMILIA")
    mytabley.Open "select * from BELCORP", cn, adOpenStatic, adLockOptimistic
    Do

        If mytabley.EOF Then Exit Do
        mytablex.Open "select * from producto where producto='" & Trim("" & mytabley.Fields("producto")) & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.AddNew
            pone_david mytablex, mytabley
            mytablex.Update
        Else
            pone_david mytablex, mytabley
            mytablex.Update

        End If

        mytablex.Close
        sdx = sdx + 1
        vr = DoEvents()
        dd = "" & sdx
        mytabley.MoveNext
    Loop
    mytabley.Close
    MsgBox "Producto proceso Terminado", 48, "Aviso"

End Sub

Sub pone_david(mytablex As ADODB.Recordset, mytabley As ADODB.Recordset)

    Dim mytablea As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    mytablex.Fields("producto") = Trim("" & mytabley.Fields("producto"))
    mytablex.Fields("barras") = Trim("" & mytabley.Fields("cod_barra"))
    mytablex.Fields("descripcio") = Mid$(Trim("" & mytabley.Fields("descripcioN")), 1, 80)
    mytablex.Fields("descorto") = Mid$(Trim("" & mytabley.Fields("desC-cortA")), 1, 20)
    mytablex.Fields("presenta") = ""
    mytablex.Fields("dsctoref") = 0
    mytablex.Fields("familia") = Mid$(Trim("" & mytabley.Fields("familia")), 1, 6)
    mytablex.Fields("subfamilia") = ""
    mytablex.Fields("vecaja") = "S"
    mytablex.Fields("igv") = Val("" & mytabley.Fields("igv"))
    mytablex.Fields("isc") = 0
    mytablex.Fields("pesokgr") = 0.001
    mytablex.Fields("comision") = 0
    mytablex.Fields("monedac") = "S"
    mytablex.Fields("unidad") = "UND"
    mytablex.Fields("factor") = 1
    mytablex.Fields("costou") = Val("" & mytabley.Fields("cu_cost_ult"))
    mytablex.Fields("costop") = Val("" & mytabley.Fields("cu_cost_ult"))
    mytablex.Fields("monedav") = "S"
    mytablex.Fields("estado") = "S"

    mytablez.Open "select * from familia where familia='" & Mid$(Trim("" & mytabley.Fields("FAMILIA")), 1, 6) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablez.RecordCount = 0 Then
        mytablez.AddNew
        mytablez.Fields("familia") = Mid$(Trim("" & mytabley.Fields("FAMILIA")), 1, 6)
        mytablez.Fields("descripcio") = Mid$(Trim("" & mytabley.Fields("FAMILIA")), 1, 30)
        mytablez.Fields("vetouch") = "S"
        mytablez.Update
    Else
        mytablez.Fields("familia") = Mid$(Trim("" & mytabley.Fields("FAMILIA")), 1, 6)
        mytablez.Fields("descripcio") = Mid$(Trim("" & mytabley.Fields("FAMILIA")), 1, 30)
        mytablez.Fields("vetouch") = "S"
        mytablez.Update

    End If

    mytablez.Close

    mytablea.Open "select * from precios where producto='" & Trim("" & mytabley.Fields("producto")) & "' and local='01'", cn, adOpenStatic, adLockOptimistic

    If mytablea.RecordCount = 0 Then
        mytablea.AddNew
        pone_david01 mytablea, mytabley, "01"
        mytablea.Update
    Else
        pone_david01 mytablea, mytabley, "01"
        mytablea.Update

    End If

    mytablea.Close

End Sub

Sub pone_david01(mytablex As ADODB.Recordset, mytabley As ADODB.Recordset, buf As String)
    mytablex.Fields("local") = buf
    mytablex.Fields("producto") = Trim("" & mytabley.Fields("producto"))
    mytablex.Fields("factor1") = 1
    mytablex.Fields("unidad1") = "UND"
    mytablex.Fields("pventa1") = Val("" & mytabley.Fields("l01_pvta"))

End Sub

Sub pasa_productos()

    Dim buf      As String

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    cn.Execute ("DELETE FROM PRODUCTOBORRAR")
    cn.Execute ("DELETE FROM preciosBORRAR")
    I = 1500
    mytablex.Open "select * from BOCADITO ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        graba_productoborra mytablex, I
        graba_preciosborra mytablex, I
        I = I + 1
        mytablex.MoveNext
    Loop
    mytablex.Close
    MsgBox "acabe", 48, "aVISO"

End Sub

Sub graba_productoborra(mytabley As ADODB.Recordset, I As Integer)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from productoborrar where producto='T" & "" & I & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.AddNew
        mytablex.Fields("producto") = "T" & "" & I
        mytablex.Fields("descripcio") = UCase$(Trim(Mid$(Trim("" & mytabley.Fields("descripcio")), 1, 40)) & "" & Trim(Mid$(Trim("" & mytabley.Fields("NOMBRE")), 1, 40)))
        mytablex.Fields("descorto") = UCase$(Mid$(Trim("" & mytabley.Fields("NOMBRE")), 1, 20))
        mytablex.Fields("presenta") = ""
        mytablex.Fields("seccion") = "PT"
        mytablex.Fields("dsctoref") = 0
        mytablex.Fields("familia") = Mid$(Trim("" & mytabley.Fields("familia")), 1, 6)
        mytablex.Fields("subfamilia") = Mid$(Trim("" & mytabley.Fields("SUBfam")), 1, 6)
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
        mytablex.Fields("estado") = "S"
        mytablex.Update

    End If

    mytablex.Close

End Sub

Sub graba_preciosborra(mytabley As ADODB.Recordset, I As Integer)

    Dim mytablex As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    mytablex.Open "select * from preciosborrar where local='00' and producto='T" & "" & I & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.AddNew
        mytablex.Fields("producto") = "T" & "" & I
        mytablex.Fields("local") = "00"
        mytablex.Fields("factor1") = 1
        mytablex.Fields("unidad1") = "UND"
        mytablex.Fields("pventa1") = Val("" & mytabley.Fields("UND1"))

        mytablex.Fields("factor2") = 25
        mytablex.Fields("unidad2") = "CAJA25"
        mytablex.Fields("pventa2") = Val("" & mytabley.Fields("UND25"))

        mytablex.Fields("factor3") = 50
        mytablex.Fields("unidad3") = "CAJA50"
        mytablex.Fields("pventa3") = Val("" & mytabley.Fields("UND50"))

        mytablex.Fields("factor4") = 75
        mytablex.Fields("unidad4") = "CAJA75"
        mytablex.Fields("pventa4") = Val("" & mytabley.Fields("UND75"))

        mytablex.Fields("factor5") = 100
        mytablex.Fields("unidad5") = "CAJ100"
        mytablex.Fields("pventa5") = Val("" & mytabley.Fields("UND100"))

        mytablex.Update

    End If

    mytablex.Close

End Sub

Sub CLIENTESMIA()

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    cn.Execute ("delete from clientes")
    cn.Execute ("delete from PROVEEDO")
    mytabley.Open "select * from clienteMIA ", cn, adOpenStatic, adLockOptimistic

    Do

        If mytabley.EOF Then Exit Do
        If UCase$(Trim("" & mytabley.Fields("tipo"))) = "PROVEEDOR" Then
            mytablex.Open "select * from proveedo where codigo='" & Mid$(Trim("" & mytabley.Fields("codigo")), 1, 11) & "'", cn, adOpenStatic, adLockOptimistic

            If mytablex.RecordCount = 0 Then
                mytablex.AddNew
                mytablex.Fields("codigo") = Mid$(Trim("" & mytabley.Fields("codigo")), 1, 11)
                mytablex.Fields("nombre") = Mid$(Trim("" & mytabley.Fields("RAZONSOCIAL")), 1, 60)
                mytablex.Fields("direccion") = Mid$(Trim("" & mytabley.Fields("direccion")), 1, 60)
                mytablex.Fields("NOMBREC") = Mid$(Trim("" & mytabley.Fields("NOMBRE")), 1, 60)
                mytablex.Fields("moneda") = "S"

                If Len(Trim("" & mytabley.Fields("codigo"))) = 11 Then
                    mytablex.Fields("tipo") = "J"

                End If

                If Len(Trim("" & mytabley.Fields("codigo"))) = 8 Then
                    mytablex.Fields("tipo") = "D"

                End If

                If Len(Trim("" & mytabley.Fields("codigo"))) < 8 Then
                    mytablex.Fields("tipo") = "O"

                End If

                mytablex.Update

            End If

            mytablex.Close

        End If

        If UCase$(Trim("" & mytabley.Fields("tipo"))) = "cliente" Then
            mytablex.Open "select * from clientes where codigo='" & Mid$(Trim("" & mytabley.Fields("codigo")), 1, 11) & "'", cn, adOpenStatic, adLockOptimistic

            If mytablex.RecordCount = 0 Then
                mytablex.AddNew
                mytablex.Fields("codigo") = Mid$(Trim("" & mytabley.Fields("codigo")), 1, 11)
                mytablex.Fields("nombre") = Mid$(Trim("" & mytabley.Fields("RAZONSOCIAL")), 1, 60)
                mytablex.Fields("direccion") = Mid$(Trim("" & mytabley.Fields("direccion")), 1, 60)
                mytablex.Fields("NOMBREC") = Mid$(Trim("" & mytabley.Fields("NOMBRE")), 1, 60)
                mytablex.Fields("moneda") = "S"

                If Len(Trim("" & mytabley.Fields("codigo"))) = 11 Then
                    mytablex.Fields("tipo") = "J"

                End If

                If Len(Trim("" & mytabley.Fields("codigo"))) = 8 Then
                    mytablex.Fields("tipo") = "D"

                End If

                If Len(Trim("" & mytabley.Fields("codigo"))) < 8 Then
                    mytablex.Fields("tipo") = "O"

                End If

                mytablex.Update

            End If

            mytablex.Close

        End If

        mytabley.MoveNext
    Loop
    '------------------------------------- ------------

    MsgBox "proceso Terminado", 48, "Aviso"

End Sub

