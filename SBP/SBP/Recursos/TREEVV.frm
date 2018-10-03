VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form treevv 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ventas"
   ClientHeight    =   8130
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Height          =   705
      Left            =   120
      Picture         =   "TREEVV.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   7140
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   12594
      _Version        =   393217
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
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
            Picture         =   "TREEVV.frx":08CA
            Key             =   "picture1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TREEVV.frx":0E64
            Key             =   "picture2"
         EndProperty
      EndProperty
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
      Height          =   840
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12975
   End
   Begin VB.Menu d89 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "treevv"
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
    treevv.Hide
    Unload treevv

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

    Dim sp5      As String

    Dim sh5      As String

    Dim sp6      As String

    Dim sh6      As String

    Dim sp7      As String

    Dim sh7      As String

    Dim sw       As Integer

    Dim found    As Integer

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    gofpago = "fpagov"
    sp = "sp"
    sp1 = "sp1"
    sp2 = "sp2"
    sp3 = "sp3"
    sp4 = "sp4"
    sp5 = "sp5"
    sp6 = "sp6"
    
    sp7 = "sp7"
    
    TreeView1.ImageList = ImageList1
    TreeView1.Nodes.Add , , sp, "Tablas", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Locales", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Almacenes", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Tipo Documentos", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Clientes", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Transportistas", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Productos", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Forma Pago", "picture1"
    
    TreeView1.Nodes.Add , , sp1, "Movimientos", "picture1"
    TreeView1.Nodes.Add sp1, tvwChild, sh1, "Guia Remision", "picture1"
    TreeView1.Nodes.Add sp1, tvwChild, sh1, "Factura Venta", "picture1"
    TreeView1.Nodes.Add sp1, tvwChild, sh1, "Nota Credito", "picture1"
    TreeView1.Nodes.Add sp1, tvwChild, sh1, "Nota Debito", "picture1"
    
    TreeView1.Nodes.Add , , sp3, "Procesos", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Cotizaciones", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Pedidos", "picture1"
    
    '''27/07/2017 kenyo Testing Completo al Sistema
    'TreeView1.Nodes.Add sp3, tvwChild, sh3, "Servicio Tecnico", "picture1"
    '''27/07/2017 kenyo Testing Completo al Sistema
       
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Comisiones Factura", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Comisiones Productos", "picture1"
    
    TreeView1.Nodes.Add , , sp2, "Reportes", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Cotizaciones ", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Cotizaciones Productos", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Pedidos ", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Pedidos Productos", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Guia Remision ", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Guia Remision Productos", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Factura Venta ", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Factura Venta Productos", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Registro Venta ", "picture1"
    
    '''27/07/2017 kenyo Testing Completo al Sistema
    'TreeView1.Nodes.Add sp2, tvwChild, sh2, "Reporte Percepcion ", "picture1"
    '''27/07/2017 kenyo Testing Completo al Sistema
       
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Lista Clientes", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Lista Precios", "picture1"
    
    '''27/07/2017 kenyo Testing Completo al Sistema
    'TreeView1.Nodes.Add sp2, tvwChild, sh2, "Documentos Emitidos Cuentas corrientes", "picture1"

    '''27/07/2017 kenyo Testing Completo al Sistema
      
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Formas Pago", "picture1"
    'TreeView1.Nodes.Add sp2, tvwChild, sh2, "Cuentas por Cobrar", "picture1"
    'TreeView1.Nodes.Add sp2, tvwChild, sh2, "Letras por Cobrar", "picture1"
    
    TreeView1.Nodes.Add , , sp5, "Esdisticas", "picture1"
    TreeView1.Nodes.Add sp5, tvwChild, sh5, "Ventas ", "picture1"
    TreeView1.Nodes.Add sp5, tvwChild, sh5, "Ventas Mensuales", "picture1"
    TreeView1.Nodes.Add sp5, tvwChild, sh5, "Ranking Productos", "picture1"
    TreeView1.Nodes.Add sp5, tvwChild, sh5, "Graficos", "picture1"
    TreeView1.Nodes.Add sp5, tvwChild, sh5, "Pedidos Productos Ranking", "picture1"
    TreeView1.Nodes.Add sp5, tvwChild, sh5, "Guia Remision Productos Ranking", "picture1"
    TreeView1.Nodes.Add sp5, tvwChild, sh5, "Factura Productos Ranking", "picture1"
    
    TreeView1.Nodes.Add , , sp4, "Tickets", "picture1"
    TreeView1.Nodes.Add sp4, tvwChild, sh4, "Cuadre Caja", "picture1"
    TreeView1.Nodes.Add sp4, tvwChild, sh4, "Unidades Vendidas", "picture1"
    TreeView1.Nodes.Add sp4, tvwChild, sh4, "Documentos Emitidos", "picture1"
    TreeView1.Nodes.Add sp4, tvwChild, sh4, "Productos Vs Documentos", "picture1"
    TreeView1.Nodes.Add sp4, tvwChild, sh4, "Ingresos Egresos Seccion", "picture1"
    TreeView1.Nodes.Add sp4, tvwChild, sh4, "Copia Cierre Caja", "picture1"
    
    TreeView1.Nodes.Add , , sp7, "Generador", "picture1"
    TreeView1.Nodes.Add sp7, tvwChild, sh7, "Documentos", "picture1"
    TreeView1.Nodes.Add sp7, tvwChild, sh7, "Productos Generador", "picture1"
    
    For I = 1 To 50
        buffer(I) = ""
    Next I

    '''27/07/2017 kenyo Testing Completo al Sistema
    'TreeView1.Nodes.Add , , sp6, "ReportesUsuariooo", "picture1"
        
    '    jindx = 0
    '    If mytablex.State = 1 Then mytablex.Close
    '   mytablex.Open "select * from archivo where menu='VENTAS' and   estado='S'", cn, adOpenStatic, adLockOptimistic
    '   If mytablex.RecordCount > 0 Then
    '        Do
    '        If mytablex.EOF Then Exit Do
    '        jindx = jindx + 1
    '        buffer(jindx) = Trim("" & mytablex.Fields("descripcio"))
    '        TreeView1.Nodes.Add sp6, tvwChild, sh6, Trim("" & mytablex.Fields("descripcio")), "picture1"
    '        mytablex.MoveNext
    '        Loop
    '   End If
    '   mytablex.Close
    
    '''27/07/2017 kenyo Testing Completo al Sistema
    
    'For i = 2 To TreeView1.Nodes.count - 1
    ' TreeView1.Nodes(i).Expanded = True
    'Next i
    
    Exit Sub

    'cmdLlenarTree_Click

End Sub
 
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

    Dim I As Integer

    If jindx > 0 Then

        For I = 1 To jindx

            If Node = buffer(I) Then
                ejecuta_reporte buffer(I)

            End If

        Next I

    End If

    If Node = "Comisiones Productos" Then
        opcion2 = "1"
        repraped.acu = "V" 'PEDIDO
        repraped.xdata = "DETALLE"
        repraped.Show 1

    End If

    If Node = "Documentos" Then
        reporgen.NAMETABLA = "Factura"
        reporgen.Show 1

    End If

    If Node = "Productos Generador" Then
        '''27/07/2017 kenyo Testing Completo al Sistema
        reporgen.NAMETABLA = "Producto"
        reporgen.Show 1

        'xprodet.Show 1
        '''27/07/2017 kenyo Testing Completo al Sistema
    End If

    If Node = "Locales" Then
        ttlocal.Show 1

    End If

    If Node = "Almacenes" Then
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

    If Node = "Transportistas" Then
        ttranspo.Show 1

    End If

    If Node = "Productos" Then
        xprodet.Show 1

    End If

    If Node = "Forma Pago" Then
        tfpago.Show 1

    End If

    If Node = "Factura Venta" Then
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        'explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
        explorap.fechai = Format(Now, "dd/mm/yyyy")
        explorap.fechaf = Format(Now, "dd/mm/yyyy")
        explorap.fechai = Format(Now, "dd/mm/yyyy")
        explorap.Caption = "Documentos Facturacion Ventas"
        explorap.tipoclie = "C"
        explorap.acu = "V"
        explorap.importacion = "COMERCIAL"

        explorap.Show 1

    End If

    If Node = "Guia Remision" Then
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        gofpago = "fpagov"
        explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
        explorap.fechaf = Format(Now, "dd/mm/yyyy")
        explorap.fechai = Format(Now, "dd/mm/yyyy")
        explorap.Caption = "Documentos Guia Remision Ventas"
        explorap.tipoclie = "C"
        explorap.acu = "T"
        explorap.Show 1

    End If

    If Node = "Nota Credito" Then
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        gofpago = "fpagov"
        'inicio 10/02/2018 pll
        'explorap.fk4844.Visible = False
        'fin 10/02/2018 pll

        ' Testing Proyecto Facturacion Electronica 05/04/2018
        explorap.DetalleSunat.Visible = True
        explorap.DarBaja.Visible = True
        ' Testing Proyecto Facturacion Electronica 05/04/2018

        explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
        explorap.fechaf = Format(Now, "dd/mm/yyyy")
        explorap.fechai = Format(Now, "dd/mm/yyyy")
        explorap.Caption = "Documentos Nota Credito Ventas"
        explorap.tipoclie = "C"
        explorap.acu = "E"
        explorap.Show 1

    End If

    If Node = "Nota Debito" Then
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        'inicio 10/02/2018 pll
        'explorap.fk4844.Visible = False
        'fin 10/02/2018 pll
        explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
        explorap.fechaf = Format(Now, "dd/mm/yyyy")
        explorap.fechai = Format(Now, "dd/mm/yyyy")
        explorap.Caption = "Documentos Nota Debito Ventas"
        explorap.tipoclie = "C"
        explorap.acu = "F"

        ' Testing Proyecto Facturacion Electronica 05/04/2018
        explorap.DetalleSunat.Visible = True
        explorap.DarBaja.Visible = True
        ' Testing Proyecto Facturacion Electronica 05/04/2018

        explorap.Show 1

    End If

    If Node = "Cotizaciones" Then
        cgusuario = "ccotizav"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "dcotizav"
        'explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
        'explorap.fechaf = Format(Now, "dd/mm/yyyy")
        explorap.Caption = "Documentos Cotizacion Ventas"
        explorap.tipoclie = "C"
        explorap.acu = "H"
        explorap.Show 1

    End If

    If Node = "Servicio Tecnico" Then
        cgusuario = "cservicio"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "dservicio"
        'explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
        'explorap.fechaf = Format(Now, "dd/mm/yyyy")
        explorap.Caption = "Servicio Tecnico"
        explorap.tipoclie = "C"
        explorap.acu = "3"
        explorap.Show 1

    End If

    If Node = "Pedidos" Then
        cgusuario = "CPEDIDOV"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DPEDIDOV"
        explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
        explorap.fechaf = Format(Now, "dd/mm/yyyy")
        explorap.Caption = "Documentos Pedidos Ventas"
        explorap.tipoclie = "C"
        explorap.acu = "I"
        explorap.Show 1

    End If

    If Node = "Guia Remision " Then
        cgusuario = "FACTURA"
        dgusuariog = "DETALLE"
        repdocum.acu = "T"
        repdocum.Show 1

    End If

    If Node = "Factura Venta " Then
        opcion2 = "0"
        cgusuario = "FACTURA"
        dgusuariog = "DETALLE"
        repdocum.Label26.Visible = True
        repdocum.comopaga.Visible = True
        repdocum.acu = "V"
        gofpago = "FPAGOV"
        repdocum.Show 1

    End If

    If Node = "Comisiones Factura" Then
        opcion2 = "4000"
        cgusuario = "FACTURA"
        dgusuariog = "DETALLE"
        repdocum.Label26.Visible = True
        repdocum.comopaga.Visible = True
        repdocum.acu = "V"
        gofpago = "FPAGOV"
        repdocum.Show 1

    End If

    If Node = "Reporte Percepcion " Then
        opcion2 = "100"
        cgusuario = "FACTURA"
        dgusuariog = "DETALLE"
        repdocum.Label26.Visible = True
        repdocum.comopaga.Visible = True
        repdocum.acu = "V"
        gofpago = "FPAGOV"
        repdocum.Show 1

    End If

    If Node = "Registro Venta " Then
        cgusuario = "FACTURA"
        dgusuariog = "DETALLE"
        repdocrv.Label24.Visible = True
        repdocrv.consolidado.Visible = True
        repdocrv.titulo = "REGISTRO DE VENTAS " & dicmoneda
        repdocrv.acu = "V"
        repdocrv.Show 1

    End If

    If Node = "Cotizaciones " Then
        cgusuario = "CcotizaV"
        dgusuariog = "DcotizaV"
        repdocum.acu = "H"
        repdocum.Show 1

    End If

    If Node = "Pedidos " Then
        cgusuario = "CPEDIDOV"
        dgusuariog = "DPEDIDOV"
        opcion2 = "P"
        repdocum.acu = "I"
        repdocum.Show 1

    End If

    If Node = "Ventas " Then
        opcion2 = "10"   'analisis de ventas
        cgusuario = "FACTURA"
        dgusuariog = "DETALLE"
        'repdocum.Label18.Visible = False
        'repdocum.Combo1.Visible = False
        repdocum.vdetalle.Enabled = False
        repdocum.vfpago.Enabled = False
        repdocum.acu = "V"
        repdocum.Show 1

    End If

    If Node = "Ventas Mensuales" Then
        opcion2 = "12"   'analisis de ventas
        cgusuario = "FACTURA"
        dgusuariog = "DETALLE"
        'repdocum.Label18.Visible = False
        'repdocum.Combo1.Visible = False
        repdocum.vdetalle.Enabled = False
        repdocum.vfpago.Enabled = False
        repdocum.acu = "V"
        repdocum.Show 1

    End If

    If Node = "Ranking Productos" Then
        opcion2 = "2"
        repraped.Label12.Visible = True
        repraped.orden.Visible = True
        repraped.acu = "V" 'PEDIDO
        repraped.xdata = "DETALLE"
        repraped.Show 1

    End If

    If Node = "Graficos" Then
        FrmChart.acu = "V"
        FrmChart.Show 1

    End If

    If Node = "Lista Clientes" Then
        trepclie.Show 1

    End If

    If Node = "Lista Precios" Then

        opcion2 = "7"
        repinv.Label27.Visible = True
        repinv.Label28.Visible = True
        repinv.Label29.Visible = True
        repinv.fechavpi.Visible = True
        repinv.fechavpf.Visible = True
        repinv.Show 1

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

    If Node = "Cuadre Caja" Then
    
        opcion2 = "1"
        opcion1 = "1"
        opcion3 = "2"
        tcuadrc1.fechai.Enabled = True
        tcuadrc1.fechaf.Enabled = True

        usuariopos = gusuario
        tcuadrc1.flagdiario = "1"
        tcuadrc1.cajero = "%"
        tcuadrc1.caja = "%"
        tcuadrc1.turno = "%"
        tcuadrc1.fechai = Format(Now, "dd/mm/yyyy")
        tcuadrc1.fechaf = Format(Now, "dd/mm/yyyy")
        tcuadrc1.horai = "01"
        tcuadrc1.horaf = "24"
        tcuadrc1.Caption = "CUADRE PARCIAL DEL DIA"
        tcuadrc1.Show 1

    End If

    If Node = "Unidades Vendidas" Then
    
        opcion1 = "3"
        opcion2 = "1"
        opcion3 = "2"
        tcuadrc1.fechai.Enabled = True
        tcuadrc1.fechaf.Enabled = True
        usuariopos = gusuario
        tcuadrc1.flagdiario = "1"
        tcuadrc1.cajero = "%"
        tcuadrc1.caja = "%"
        tcuadrc1.turno = "%"
        tcuadrc1.horai = "01"
        tcuadrc1.horaf = "24"
        tcuadrc1.fechai = Format(Now, "dd/mm/yyyy")
        tcuadrc1.fechaf = Format(Now, "dd/mm/yyyy")
        tcuadrc1.Caption = "UNIDADES VENDIDAS"
        tcuadrc1.check3d1.Visible = True
        tcuadrc1.check3d2.Visible = True
        tcuadrc1.check3d3.Visible = True
        tcuadrc1.Show 1

    End If

    If Node = "Documentos Emitidos" Then
    
        opcion1 = "2"
        opcion2 = "1"
        opcion3 = "2"
        tcuadrc1.fechai.Enabled = True
        tcuadrc1.fechaf.Enabled = True
        usuariopos = gusuario
        tcuadrc1.flagdiario = "1"

        tcuadrc1.cajero = "%"
        tcuadrc1.caja = "%"
        tcuadrc1.turno = "%"
        tcuadrc1.horai = "01"
        tcuadrc1.horaf = "24"
        tcuadrc1.fechai = Format(Now, "dd/mm/yyyy")
        tcuadrc1.fechaf = Format(Now, "dd/mm/yyyy")
        tcuadrc1.Caption = "DOCUMENTOS EMITIDOS"
        tcuadrc1.check3d1.Visible = True
        tcuadrc1.check3d2.Visible = True
        tcuadrc1.check3d3.Visible = True
        tcuadrc1.Show 1

    End If

    If Node = "Productos Vs Documentos" Then
    
        opcion1 = "4"
        opcion2 = "1"
        opcion3 = "2"
        tcuadrc1.fechai.Enabled = True
        tcuadrc1.fechaf.Enabled = True
        tcuadrc1.cajero = "%"
        tcuadrc1.caja = "%"
        tcuadrc1.turno = "%"
        tcuadrc1.horai = "01"
        tcuadrc1.horaf = "24"
        tcuadrc1.flagdiario = "1"

        tcuadrc1.fechai = Format(Now, "dd/mm/yyyy")
        tcuadrc1.fechaf = Format(Now, "dd/mm/yyyy")
        tcuadrc1.Caption = "PRODUCTOS VS DOCUMENTOS"
        tcuadrc1.check3d1.Visible = True
        tcuadrc1.check3d2.Visible = True
        tcuadrc1.check3d3.Visible = True
        tcuadrc1.Show 1

    End If

    If Node = "Ingresos Egresos Seccion" Then
    
        opcion1 = "20"
        opcion2 = "2"
        opcion3 = ""
        tcuadrc1.flagdiario = "1"
        tcuadrc1.fechai.Enabled = True
        tcuadrc1.fechaf.Enabled = True
        usuariopos = gusuario
        tcuadrc1.cajero = "%"
        tcuadrc1.caja = "%"
        tcuadrc1.turno = "%"
        tcuadrc1.horai = "01"
        tcuadrc1.horaf = "24"
        tcuadrc1.fechai = Format(Now, "dd/mm/yyyy")
        tcuadrc1.fechaf = Format(Now, "dd/mm/yyyy")
        tcuadrc1.Caption = "DOCUMENTOS EMITIDOS PERIODICO"
        tcuadrc1.check3d1.Visible = True
        tcuadrc1.check3d2.Visible = True
        tcuadrc1.check3d3.Visible = True
        tcuadrc1.Show 1
        tcuadrc1.flagdiario = ""

    End If

    If Node = "Copia Cierre Caja" Then

        opcion1 = "5"
        opcion2 = "1"
        opcion3 = "2"
    
        tcuadrc1.fechai.Enabled = True
        tcuadrc1.fechaf.Enabled = True
    
        usuariopos = gusuario
        tcuadrc1.tipoexterno.Visible = True
        tcuadrc1.numcuadre.Visible = True
        'tcuadrc1.flagdiario = "1"
        tcuadrc1.cajero = "%"
        tcuadrc1.caja = "%"
        tcuadrc1.turno = "%"
        tcuadrc1.fechai = Format(Now, "dd/mm/yyyy")
        tcuadrc1.fechaf = Format(Now, "dd/mm/yyyy")
        tcuadrc1.horai = "01"
        tcuadrc1.horaf = "24"
        tcuadrc1.Caption = "COPIA CIERRE DEL DIA"
        tcuadrc1.Show 1

    End If

    If Node = "Documentos Emitidos Cuentas corrientes" Then
        opcion2 = "900"
        cgusuario = "FACTURA"
        dgusuariog = "DETALLE"
        repdocum.Label26.Visible = True
        repdocum.comopaga.Visible = True
        repdocum.titulo = "Estado Facturas"
        repdocum.acu = "V"
        gofpago = "FPAGOV"
        repdocum.Show 1

    End If

    If Node = "Formas Pago" Then
        repfpago.Show 1

    End If

    If Node = "Cotizaciones Productos" Then
        opcion2 = "1"
        repraped.acu = "H" 'PEDIDO
        repraped.xdata = "dcotizav"
        repraped.Show 1

    End If

    If Node = "Pedidos Productos" Then
        opcion2 = "1"
        repraped.acu = "I" 'PEDIDO
        repraped.xdata = "dpedidov"
        repraped.Show 1

    End If

    If Node = "Pedidos Productos Ranking" Then
        opcion2 = "2"
        repraped.Label12.Visible = True
        repraped.orden.Visible = True
        repraped.acu = "" 'PEDIDO
        repraped.xdata = "Dpedidov"
        repraped.Show 1

    End If

    If Node = "Guia Remision Productos" Then
        opcion2 = "1"
        repraped.acu = "T" 'PEDIDO
        repraped.xdata = "detalle"
        repraped.Show 1

    End If

    If Node = "Guia Remision Productos Ranking" Then
        opcion2 = "2"
        repraped.Label12.Visible = True
        repraped.orden.Visible = True
        repraped.acu = "H" 'PEDIDO
        repraped.xdata = "detalle"
        repraped.Show 1

    End If

    If Node = "Factura Venta Productos" Then
        opcion2 = "1"
        repraped.acu = "V" 'PEDIDO
        repraped.xdata = "DETALLE"
        repraped.Show 1

    End If

    If Node = "Factura Productos Ranking" Then
        opcion2 = "2"
        repraped.Label12.Visible = True
        repraped.orden.Visible = True
        repraped.acu = "V" 'PEDIDO
        repraped.xdata = "Detalle"
        repraped.Show 1

    End If

    If Node = "Cuentas por Cobrar" Then

        repctaxc.acu = "V"
        repctaxc.Show 1

    End If

    If Node = "Letras por Cobrar" Then
        REPLETRA.titulo = "Letras por Cobrar"
        REPLETRA.acu = "V"
        REPLETRA.Show 1

    End If

End Sub

Sub ejecuta_reporte(buf As String)
    reporgen.NAMETABLA = "ventasProducto"
    reporgen.Show 1

End Sub
