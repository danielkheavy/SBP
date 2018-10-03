VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form treev 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inventario"
   ClientHeight    =   9735
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   10485
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5880
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "treev.frx":0000
            Key             =   "picture1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "treev.frx":059A
            Key             =   "picture2"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton btnsalir 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   120
      Picture         =   "treev.frx":0B34
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir todo"
      Top             =   120
      Width           =   1695
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   14843
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label registros 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10560
      TabIndex        =   3
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12975
   End
   Begin VB.Menu d89 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "treev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim buffer(50) As String

Dim jindx      As Integer

Option Explicit

Private Sub btnsalir_Click()
    d89_Click

End Sub

Private Sub d89_Click()
    treev.Hide
    Unload treev

End Sub

Private Sub Form_Load()

    Dim sp       As String

    Dim sh       As String

    Dim sp1      As String

    Dim sh1      As String

    Dim sp2      As String

    Dim sh2      As String

    Dim sp3      As String

    Dim sh3      As String

    Dim sp4      As String

    Dim sh4      As String

    Dim mytablex As New ADODB.Recordset

    Dim I        As Integer

    sp = "sp"
    sp1 = "sp1"
    sp2 = "sp2"
    sp3 = "sp3"
    sp4 = "sp4"
    
    'With TreeView1
    '.Style = tvwTreelinesPlusMinusText
    '.LineStyle = tvwRootLines
    '.PathSeparator = "\"
    '.Indentation = Screen.TwipsPerPixelX * 5 '256
    '.LabelEdit = tvwManual
    '.SingleSel = False
    '.HideSelection = False
    '.Refresh
    'End With
    
    'For i = 1 To 1
    '  ImageList1.ListImages.Add , , Picture1
    'Next i
    TreeView1.ImageList = ImageList1
    TreeView1.Nodes.Add , , sp, "Tablas", "picture1"
    
    TreeView1.Nodes.Add sp, tvwChild, sh, "Locales", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Almacenes", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Tipo Documentos", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Clientes", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Proveedor", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Transportistas", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Productos", "picture1"
    
    TreeView1.Nodes.Add , , sp1, "Movimientos", "picture1"
    TreeView1.Nodes.Add sp1, tvwChild, sh1, "Guias de Ingreso", "picture1"
    TreeView1.Nodes.Add sp1, tvwChild, sh1, "Guias de Salida", "picture1"
    TreeView1.Nodes.Add sp1, tvwChild, sh1, "Traslados", "picture1"
    TreeView1.Nodes.Add sp1, tvwChild, sh1, "SaldoInicial", "picture1"
     
    'TreeView1.Nodes.Add sp1, tvwChild, sh1, "Traslados Rapidos", "picture2"
    TreeView1.Nodes.Add sp1, tvwChild, sh1, "ConteoFisico", "picture1"
    'TreeView1.Nodes.Add sp1, tvwChild, sh1, "Saldo Inicial", "picture1"
    'TreeView1.Nodes.Add sp1, tvwChild, sh1, "Saldo Inicial", "picture1"
    
    'TreeView1.Nodes.Add sp1, tvwChild, sh1, "Regularizacion", "picture1"
    
    TreeView1.Nodes.Add , , sp3, "Procesos", "picture1"
    'TreeView1.Nodes.Add sp3, tvwChild, sh3, "Pedidos Automaticos", "picture2"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Requerimientos de Almacen", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Recalculo Saldos", "picture1"
    'TreeView1.Nodes.Add sp3, tvwChild, sh3, "Cierre Inventario", "picture1"
    'TreeView1.Nodes.Add sp3, tvwChild, sh3, "Generacion Etiquetas", "picture1"
    
    TreeView1.Nodes.Add , , sp2, "Reportes", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Saldo Inicial ", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Saldo Actual ", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Kardex", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Kardex Sunat", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Saldos Almacenes", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Entradas Salidas ", "picture1"
    
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Analisis Lineas Tallas", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Pedidos Almacenes", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Saldos a un Periodo", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Productos Sin Rotacion", "picture1"
    
    '''27/07/2017 kenyo Testing Completo al Sistema
    'TreeView1.Nodes.Add sp2, tvwChild, sh2, "Productos Recetas", "picture1"
    '''27/07/2017 kenyo Testing Completo al Sistema
    
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Lista para Conteo", "picture1"
    'TreeView1.Nodes.Add sp2, tvwChild, sh2, "AnalisisRequerimiento", "picture1"
    
    'TreeView1.Nodes.Add sp2, tvwChild, sh2, "KardexFormato", "picture1"
    
    For I = 1 To 50
        buffer(I) = ""
    Next I
    
    '''27/07/2017 kenyo Testing Completo al Sistema
    'TreeView1.Nodes.Add , , sp4, "ReportesUsuario", "picture1"
    '''27/07/2017 kenyo Testing Completo al Sistema
    '------------------
    
    jindx = 0

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from archivo where menu='ALMACEN' and   estado='S'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        Do

            If mytablex.EOF Then Exit Do
            jindx = jindx + 1
            buffer(jindx) = Trim("" & mytablex.Fields("descripcio"))
            TreeView1.Nodes.Add sp4, tvwChild, sh4, Trim("" & mytablex.Fields("descripcio")), "picture1"
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

    For I = 2 To TreeView1.Nodes.count - 1
        'TreeView1.Nodes(i).ExpandedImage = "Open"
        TreeView1.Nodes(I).Expanded = True
    Next I

    Exit Sub
    'cmdLlenarTree_Click

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

    Dim I   As Integer

    Dim buf As String

    'If jindx > 0 Then
    'For i = 1 To jindx
    '    If Node = buffer(i) Then
    '       ejecuta_reporte buffer(i)
    '    End If
    'Next i
    'End If

    If Node = "KardexFormato" Then
        cn.Execute ("delete from kardex ")
        buf = "INSERT INTO kardex "
        buf = buf & " (producto, descripcio, unidad, factor, tipo, serie, numero, fecha, entrada,salida)"
        buf = buf & "SELECT     PRODUCTO, DESCRIPCIO, UNIDAD, FACTOR, TIPO, SERIE, NUMERO, FECHA, CANTIDAD*factor,0"
        buf = buf & " From detalle  "
        buf = buf & " WHERE     (ACU = 'J') OR  (ACU = 'K') OR (ACU = 'L') OR  (ACU = 'M') OR  (ACU = 'S') and estado='2'"
        cn.Execute (buf)

        buf = "INSERT INTO kardex "
        buf = buf & " (producto, descripcio, unidad, factor, tipo, serie, numero, fecha, entrada,salida)"
        buf = buf & "SELECT     PRODUCTO, DESCRIPCIO, UNIDAD, FACTOR, TIPO, SERIE, NUMERO, FECHA, 0,CANTIDAD*factor"
        buf = buf & " From detalle "
        buf = buf & " WHERE     (ACU = 'A') OR  (ACU = 'B') OR (ACU = 'C') OR  (ACU = 'D') OR  (ACU = 'T') and estado='2'"
        cn.Execute (buf)

        cn.Execute ("update kardex set saldo=0")
        cn.Execute ("delete from kardexsaldo")
        buf = " insert into kardexsaldo "
        buf = buf & " (producto,saldo) select producto,sum(entrada-salida) from kardex group by producto "
        cn.Execute (buf)

        'Do
        'If mytablex.EOF Then Exit Do
        'Loop

        ejecuta_kardex

        'reporgen.NAMETABLA = "kardex"
        'reporgen.Show 1
    End If

    If Node = "SaldoInicial" Then

        cgusuario = "csaldoini"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "dsaldoini"
        explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
        explorap.fechaf = Format(Now, "dd/mm/yyyy")
        explorap.fechai = Format(Now, "dd/mm/yyyy")
        explorap.Caption = "Documento Ingreso Saldo Inicial"
        explorap.tipoclie = "P"
        explorap.acu = "S"
        explorap.Show 1

    End If

    If Node = "Locales" Then
        ttlocal.Show 1

    End If

    If Node = "Almacenes" Then
        'talmacen.fdk8923.Visible = True
        'talmacen.ajdu1.Visible = False
        'talmacen.f8443.Visible = False
        'talmacen.bo712.Visible = False
        'talmacen.cmdAddEntry.Visible = False
        'talmacen.cmdSave.Visible = False
        'talmacen.cmdDelete.Visible = False
        talmacen.Show 1

    End If

    If Node = "Tipo Documentos" Then
        tdocumen.Show 1

    End If

    If Node = "Clientes" Then
        tnclie.DBPROV = "clientes"
        tnclie.Show 1

    End If

    If Node = "Proveedor" Then
        tnclie.DBPROV = "Proveedo"
        tnclie.Show 1

    End If

    If Node = "Traslados" Then
        cgusuario = "ctraslad"
        'cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "dtraslad"
        'dgusuariog = "DETALLE"
        'explorap.fk4844.Visible = False 'inicio 09/02/2018 pll
        explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
        explorap.fechaf = Format(Now, "dd/mm/yyyy")
        explorap.fechai = Format(Now, "dd/mm/yyyy")
        explorap.Caption = "Traslado entre almacen de un mismo establecimiento"
        explorap.tipoclie = "V"
        explorap.acu = "Z"
        explorap.Show 1

    End If

    If Node = "Transportistas" Then
        ttranspo.Show 1

    End If

    If Node = "Guias de Ingreso" Then
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        'explorap.fk4844.Visible = False 'inicio 09/02/2018 pll

        explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
        explorap.fechaf = Format(Now, "dd/mm/yyyy")
        explorap.fechai = Format(Now, "dd/mm/yyyy")

        explorap.Caption = "Documentos Guia Remision Compra"
        explorap.tipoclie = "V"
        explorap.acu = "S"
        explorap.tinterno = "S"
        explorap.Show 1

    End If

    If Node = "Guias de Salida" Then
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        'explorap.fk4844.Visible = False 'inicio 09/02/2018 pll
        explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
        explorap.fechaf = Format(Now, "dd/mm/yyyy")
        explorap.fechai = Format(Now, "dd/mm/yyyy")
        explorap.Caption = "Documentos Guia Remision Ventas"
        explorap.tipoclie = "V"
        explorap.tinterno = "S"
        explorap.acu = "T"
        explorap.Show 1

    End If

    If Node = "Productos" Then
        xprodet.Show 1

    End If

    If Node = "Regularizacion" Then
        tsiconte.Show 1

    End If

    'If Node = "Saldo Inicial" Then
    '   tsaldoin.Show 1
    'End If
    If Node = "Saldo Inicial" Then
        talmacen.fdk8923.Visible = True
        talmacen.ajdu1.Visible = False
        talmacen.f8443.Visible = False
        talmacen.bo712.Visible = False
        talmacen.cmdAddEntry.Visible = False
        talmacen.cmdSave.Visible = False
        talmacen.cmdDelete.Visible = False
        'tinicont.Show 1
        talmacen.Show 1

    End If

    If Node = "Recalculo Saldos" Then
        trecalcu.Show 1

    End If

    If Node = "ConteoFisico" Then
        tpefisic.fis8912.Visible = True
        '' 03/07/2018 Conteo Fisico Sistema
        '   tpefisic.ajdu1.Visible = True
        '   tpefisic.f8443.Visible = False
        '   tpefisic.bo712.Visible = False
        '   tpefisic.cmdAddEntry.Visible = False
        '   tpefisic.cmdSave.Visible = False
        '   tpefisic.cmdDelete.Visible = False
        tpefisic.ajdu1.Visible = True
        tpefisic.f8443.Visible = True
        tpefisic.bo712.Visible = True
        tpefisic.cmdAddEntry.Visible = True
        tpefisic.cmdSave.Visible = True
        tpefisic.cmdDelete.Visible = True
        '' 03/07/2018 Conteo Fisico Sistema
   
        tpefisic.Show 1

        'tconoffc.Show 1
    End If

    If Node = "Generacion Etiquetas" Then
        tcxbarra.Show 1

    End If

    If Node = "Kardex" Then
        opcion2 = "1"
        repinv.excell.Visible = True

        repinv.Label15.Visible = True
        repinv.Label16.Visible = True
        repinv.fechai.Visible = True
        repinv.fechaf.Visible = True
        repinv.fechai.Enabled = True

        '''10/08/2017 kenyo Mejor Kardex Producto
        repinv.quecosto.Visible = True
        repinv.Label33.Visible = True

        '''10/08/2017 kenyo Mejor Kardex Producto

        repinv.Show 1

    End If

    If Node = "Kardex Sunat" Then
        opcion2 = "100"
        repinv.excell.Visible = True
   
        repinv.Label15.Visible = True
        repinv.Label16.Visible = True
        repinv.fechai.Visible = True
        repinv.fechaf.Visible = True
        repinv.fechai.Enabled = True
        repinv.Show 1

    End If

    If Node = "Requerimientos de Almacen" Then
        cgusuario = "cREQUISA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "dREQUISA"
        explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
        explorap.fechaf = Format(Now, "dd/mm/yyyy")
        explorap.Caption = "Requerimientos de Almacen"
        explorap.acu = "Q"
        explorap.tipoclie = "V"

        explorap.Show 1

    End If

    If Node = "Lista para Conteo" Then

        opcion2 = "72"
        repinv.Label27.Visible = True
        repinv.Label28.Visible = True
        repinv.Label29.Visible = True
        repinv.fechavpi.Visible = True
        repinv.fechavpf.Visible = True
        repinv.Show 1

    End If

    If Node = "Saldo Inicial " Then
        opcion2 = "2"
        repinv.excell.Visible = True
        repinv.Label23.Visible = True
        repinv.conteo.Visible = True
        repinv.Label15.Visible = True
        repinv.fechai.Visible = True
        'repinv.fechaf.Visible = True
        'repinv.Label17.Visible = True
        'repinv.Label16.Visible = True
        repinv.Combo1.Visible = True
        repinv.xbasedatos = "dsaldoini"
        repinv.Show 1

    End If

    If Node = "Saldo Actual " Then
        opcion2 = "4"
        repinv.Label33.Visible = True
        repinv.quecosto.Visible = True

        repinv.excell.Visible = True
        repinv.Label17.Visible = True
        repinv.Combo1.Visible = True
        repinv.Label25.Visible = True
        repinv.gcanti.Visible = True
        repinv.Show 1

    End If

    If Node = "Entradas Salidas " Then
        opcion2 = "44"
        repinv.excell.Visible = True

        repinv.Label15.Visible = True
        repinv.Label16.Visible = True
        repinv.fechai.Visible = True
        repinv.fechaf.Visible = True
        repinv.fechai.Enabled = True
        repinv.Show 1

    End If

    If Node = "Saldos Almacenes" Then
        opcion2 = "6"
        repinv.bodega.Enabled = False
        repinv.Show 1

    End If

    If Node = "AnalisisRequerimiento" Then
        opcion2 = "16"
        repinv.bodega.Enabled = False
        repinv.Show 1

    End If

    If Node = "Analisis Lineas Tallas" Then
        opcion2 = "3"
        repinv.Label17.Visible = True
        repinv.Combo1.Visible = True
        repinv.Show 1

    End If

    If Node = "Pedidos Almacenes" Then
        opcion2 = "3"
        repraped.Label12.Visible = True
        repraped.orden.Visible = True
        repraped.acu = "Q" 'PEDIDO
        repraped.xdata = "DREQUISA"
        repraped.Show 1

    End If

    If Node = "Saldos a un Periodo" Then
        opcion2 = "8"
        repinv.Label15.Visible = True
        repinv.fechai.Visible = True
        repinv.fechaf.Visible = True
        repinv.Label16.Visible = True
        repinv.Label16 = "FechaFin"

        repinv.excell.Visible = True
        'repinv.Label17.Visible = True
        'repinv.Combo1.Visible = True
        repinv.Show 1

    End If

    If Node = "Pedidos Automaticos" Then

        Dim found As Integer

        found = copiar_temporalpe()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
            End
            Exit Sub

        End If

        dgusuario = "_t" & gusuario
        tpedloca.Show 1

    End If

    If Node = "Traslados Rapidos" Then

        'doctrasl.Show 1
    End If

    If Node = "Productos Sin Rotacion" Then
        opcion2 = "948"
        'repinv.excell.Visible = True
        repinv.Label17.Visible = True
        repinv.Combo1.Visible = True
        'repinv.Label25.Visible = True
        'repinv.gcanti.Visible = True

        repinv.Label30.Visible = True
        repinv.Label31.Visible = True
        repinv.Label32.Visible = True

        repinv.fechari.Visible = True
        repinv.fecharf.Visible = True

        repinv.excell.Visible = True

        'AGREGADO KENYO
        repinv.fecharf = Format(Now, "dd/mm/yyyy")
        repinv.fechari = "01" & "/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")

        repinv.Show 1

    End If

    If Node = "Productos Recetas" Then
        opcion2 = "1948"
        'repinv.excell.Visible = True
        repinv.Label17.Visible = True
        repinv.Combo1.Visible = True
        'repinv.Label25.Visible = True
        'repinv.gcanti.Visible = True

        repinv.Label30.Visible = True
        repinv.Label31.Visible = True
        repinv.Label32.Visible = True
        repinv.fechari.Visible = True
        repinv.fecharf.Visible = True
        repinv.Show 1

    End If

    TreeView1.refresh

End Sub

Sub ejecuta_kardex()

    Dim vr

    Dim I           As Long

    Dim xnro        As Double

    Dim mytablex    As New ADODB.Recordset

    Dim mytabley    As New ADODB.Recordset

    Dim movimiento  As Double

    Dim saldoprevio As Double

    Dim nuevosaldo  As Double

    Dim sdx1        As Double

    Dim sdx         As Double

    sdx1 = 0
    sdx = 0
    mytabley.Open "select * from kardexsaldo ", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount = 0 Then
        mytabley.Close
        Exit Sub

    End If

    sdx1 = mytabley.RecordCount
    Do

        If mytabley.EOF Then Exit Do
        sdx = sdx + 1
        vr = DoEvents()
        registros = "" & sdx1 & " " & sdx

        If mytablex.State = 1 Then
            mytablex.Close
            Set mytablex = Nothing

        End If

        saldoprevio = 0
        nuevosaldo = 0
        mytablex.Open "select * from kardex  where producto='" & "" & mytabley.Fields("producto") & "' order by producto,fecha", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then

            For I = 1 To mytablex.RecordCount

                If Val("" & mytablex.Fields("entrada")) > 0 Then movimiento = Val("" & mytablex.Fields("entrada")) Else movimiento = Val("" & mytablex.Fields("salida"))
                If I = 1 Then
                    mytablex.Fields("saldo").Value = movimiento + saldoprevio
                    nuevosaldo = movimiento + saldoprevio
                Else
                    mytablex.Fields("saldo").Value = movimiento + nuevosaldo
                    nuevosaldo = movimiento + nuevosaldo

                End If

                mytablex.Update
                mytablex.MoveNext
            Next I

        End If

        mytabley.MoveNext
    Loop
    'mytablex.Close

    excell_pasamos mytablex
    mytablex.Close

End Sub

Sub excell_pasamos(mysnapx As ADODB.Recordset)
    
    Dim xlApp     As Excel.Application

    Dim xlBook    As Excel.Workbook

    Dim xlSheet   As Excel.Worksheet

    Dim sFileName As String

    On Error GoTo PROC_ERR

    'MsgBox "Please format Date column to Date and Time column to time in Excel.", vbInformation, "Message"
    If mysnapx.RecordCount = 0 Then
        MsgBox "No existen Datos ", 48, "Aviso"
        Exit Sub

    End If
    
    sFileName = App.path & "\Time Log as of " & CStr(Format(Now, "mm-dd-yyyy")) & ".xls"

    ExportRecordSetToExcel mysnapx, sFileName, "", "TimeLog"

    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Open(sFileName)
    xlApp.Application.Visible = True

PROC_EXIT:
    Set mysnapx = Nothing
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Exit Sub

PROC_ERR:
    MsgBox "Primero Ejecutar: " & Err.Number & " = " & Err.Description

    Resume PROC_EXIT

End Sub
