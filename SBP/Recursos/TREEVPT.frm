VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form treevpt 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tienda"
   ClientHeight    =   9240
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   10500
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
      Height          =   825
      Left            =   120
      Picture         =   "TREEVPT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir todo"
      Top             =   120
      Width           =   1695
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   8175
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   14420
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
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
      Left            =   5760
      Top             =   3720
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
            Picture         =   "TREEVPT.frx":08CA
            Key             =   "picture1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TREEVPT.frx":0E64
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
      Height          =   1080
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12975
   End
   Begin VB.Menu d89 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "treevpt"
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
    treevpt.Hide
    Unload treevpt

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

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    sp = "sp"
    sp1 = "sp1"
    sp2 = "sp2"
    sp3 = "sp3"
    sp4 = "sp4"
    sp5 = "sp5"
    sp6 = "sp6"
    TreeView1.ImageList = ImageList1
    
    TreeView1.Nodes.Add , , sp, "Tablas", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "ConFigCorreo", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Clientes", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Parametros Caja", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Turnos", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Forma Pago", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Personal", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Edicion Formatos", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Salon", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Salon Numero/Mesas/Habit", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Servicios", "picture1"
    'TreeView1.Nodes.Add sp, tvwChild, sh, "ServicioMesa", "picture1"
    
    TreeView1.Nodes.Add sp, tvwChild, sh, "Grupo Comentario", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Comentario", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Visor Cliente", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Caja Defecto", "picture1"
    
    TreeView1.Nodes.Add , , sp3, "Procesos", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Proformas", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Caja Registradora", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Caja Registradora Touch Screen", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Caja Registradora Parqueo", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Centralizacion Cajas", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Verificador Precios", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Programa Zebra", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Cocina Monitor", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Entrega Monitor", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Cuadre Ciego", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Centralizacion Recepcion", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Centralizacion Envio", "picture1"
    
    TreeView1.Nodes.Add , , sp4, "Reportes Tickets", "picture1"
    TreeView1.Nodes.Add sp4, tvwChild, sh4, "Cuadre Caja", "picture1"
    TreeView1.Nodes.Add sp4, tvwChild, sh4, "Unidades Vendidas", "picture1"
    TreeView1.Nodes.Add sp4, tvwChild, sh4, "Documentos Emitidos", "picture1"
    TreeView1.Nodes.Add sp4, tvwChild, sh4, "Productos Vs Documentos", "picture1"
    TreeView1.Nodes.Add sp4, tvwChild, sh4, "Ingresos Egresos Seccion", "picture1"
    TreeView1.Nodes.Add sp4, tvwChild, sh4, "Copia Cierre Caja", "picture1"
    
    TreeView1.Nodes.Add , , sp2, "Reportes", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Factura Venta ", "picture1"
    
    '''27/07/2017 kenyo Testing Completo al Sistema
    'TreeView1.Nodes.Add sp2, tvwChild, sh2, "Registro Ventas  ", "picture1"
    '''27/07/2017 kenyo Testing Completo al Sistema
    
    'TreeView1.Nodes.Add sp2, tvwChild, sh2, "Factura Venta Productos", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Registro Venta ", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Lista Clientes", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Lista Precios", "picture1"
    
    '''27/07/2017 kenyo Testing Completo al Sistema
    'TreeView1.Nodes.Add sp2, tvwChild, sh2, "Documentos Emitidos Cuentas corrientes", "picture1"
    '''27/07/2017 kenyo Testing Completo al Sistema
    
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Formas Pago", "picture1"
    
    '''27/07/2017 kenyo Testing Completo al Sistema
    'TreeView1.Nodes.Add sp2, tvwChild, sh2, "Cuentas por Cobrar", "picture1"
    'TreeView1.Nodes.Add sp2, tvwChild, sh2, "Letras por Cobrar", "picture1"
    '''27/07/2017 kenyo Testing Completo al Sistema
     
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Recibos Ingreso", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Recibos Egreso", "picture1"
     
    TreeView1.Nodes.Add , , sp5, "Esdisticas", "picture1"
    TreeView1.Nodes.Add sp5, tvwChild, sh5, "Ventas ", "picture1"
    TreeView1.Nodes.Add sp5, tvwChild, sh5, "Ventas Mensuales", "picture1"
    TreeView1.Nodes.Add sp5, tvwChild, sh5, "Ranking Productos", "picture1"
    TreeView1.Nodes.Add sp5, tvwChild, sh5, "Graficos", "picture1"
   
    For I = 1 To 50
        buffer(I) = ""
    Next I
     
    '''27/07/2017 kenyo Testing Completo al Sistema
    'TreeView1.Nodes.Add , , sp6, "ReportesUsuario", "picture1", "picture1"
    '''27/07/2017 kenyo Testing Completo al Sistema
    '------------------
    
    jindx = 0

    If mytablex.State = 1 Then mytablex.Close
    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from archivo where menu='TIENDA' and   estado='S'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        Do

            If mytablex.EOF Then Exit Do
            jindx = jindx + 1
            buffer(jindx) = Trim("" & mytablex.Fields("descripcio"))
            TreeView1.Nodes.Add sp6, tvwChild, sh6, Trim("" & mytablex.Fields("descripcio")), "picture1"
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
    
    Exit Sub

    'cmdLlenarTree_Click

End Sub
 
Private Sub Image1_Click()

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

    Dim I As Integer

    'If jindx > 0 Then
    'For i = 1 To jindx
    '    If Node = buffer(i) Then
    '       ejecuta_reporte buffer(i)
    '    End If
    'Next i
    'End If
    If Node = "Caja Defecto" Then
        tcajade.Show 1

    End If

    If Node = "Entrega Monitor" Then
        tmoentre.Show 1

    End If

    'If CDate(Format(Now, "dd/mm/yyyy")) < CDate("05/01/2014") Then
    If Node = "Centralizacion Recepcion" Then
        tupload.Show 1

    End If

    If Node = "Centralizacion Envio" Then
        tload.Show 1

    End If

    'End If

    If Node = "Cuadre Ciego" Then
        tcajacie.Show 1

    End If

    If Node = "Programa Zebra" Then
        xzebra.Show 1

    End If

    If Node = "Servicios" Then
        tservice.Show 1

    End If

    If Node = "Clientes" Then
        tnclie.DBPROV = "clientes"
        tnclie.Show 1

    End If

    If Node = "Visor Cliente" Then
        tvisorc.Show 1

    End If

    If Node = "ServicioMesa" Then
        tsermesa.Show 1

    End If

    If Node = "Parametros Caja" Then
        If busca_clave1(gusuario) <> "S" Then
            MsgBox "No tiene Permiso", 48, "Aviso"
            Exit Sub

        End If

        tcaja.Show 1

    End If

    If Node = "Turnos" Then
        tturno.Show 1

    End If

    If Node = "Forma Pago" Then
        tfpago.Show 1

    End If

    If Node = "Personal" Then
        If busca_clave1(gusuario) <> "S" Then
            MsgBox "No tiene Permiso", 48, "Aviso"
            Exit Sub

        End If

        tpersona.Show 1

    End If

    If Node = "Comentario" Then
        tuchcome.Show 1

    End If

    If Node = "Edicion Formatos" Then
        If busca_clave1(gusuario) <> "S" Then
            MsgBox "No tiene Permiso", 48, "Aviso"
            Exit Sub

        End If

        teditor.Show 1

    End If

    If Node = "ConFigCorreo" Then
        tcosms.Show 1

    End If

    If Node = "Proformas" Then
        menucaja.Label3 = "USUARIO"
        menucaja.acu = "T"
        menucaja.Show 1

    End If

    If Node = "Cocina Monitor" Then
        'menucaja.Label3 = "USUARIO"
        'menucaja.acu = "T"
        kitchen.Show 1

    End If

    If Node = "Caja Registradora" Then
        menucaja.Label3 = "CAJERO"
        menucaja.acu = "C"
        menucaja.Label1.Visible = True
        menucaja.turno.Visible = True
        menucaja.Label5 = "CAJA"
        menucaja.tipoterminal = "NORMAL"
        menucaja.Show 1

    End If

    If Node = "Centralizacion Cajas" Then
        CENTRADI.Show 1

    End If

    If Node = "Verificador Precios" Then
        tncr1.local1 = "01"
        tncr1.Show 1

    End If

    If Node = "Registro Ventas  " Then
        treporte.tituloreporte = "Registro de Ventas"
        treporte.archivoreporte = globaldir & "\reportes\registroventas.rpt"
        treporte.acu = "V"
        treporte.Show 1

    End If

    If Node = "Grupo Comentario" Then
        tgrupoco.Show 1

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

    If Node = "Registro Venta " Then
        cgusuario = "FACTURA"
        dgusuariog = "DETALLE"
        repdocrv.Label24.Visible = True
        repdocrv.consolidado.Visible = True
        repdocrv.titulo = "REGISTRO DE VENTAS " & dicmoneda
        repdocrv.acu = "V"
        repdocrv.Show 1

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

    If Node = "Formas Pago" Then
        repfpago.Show 1

    End If

    If Node = "Cuentas por Cobrar" Then
        'repctaxc.acu = "V"
        'repctaxc.Show 1
        trepoctc.tituloreporte = "Cuentas Por Cobrar"
        trepoctc.archivoreporte = globaldir & "\reportes\cuentaxc.rpt"
        trepoctc.acu = "V"
        trepoctc.Show 1

    End If

    If Node = "Letras por Cobrar" Then
        REPLETRA.titulo = "Letras por Cobrar"
        REPLETRA.acu = "V"
        REPLETRA.Show 1

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

    If Node = "Factura Venta Productos" Then
        'opcion2 = "1"
        'repraped.acu = "V" 'PEDIDO
        'repraped.xdata = "DETALLE"
        'repraped.Show 1
        treporte.tituloreporte = "Reporte ventas Detalle"
        treporte.archivoreporte = globaldir & "\reportes\facturadetalle.rpt"
        treporte.acu = "V"
        treporte.Show 1

    End If

    If Node = "Recibos Ingreso" Then
        repingre.acu = "W"
        repingre.Show 1

    End If

    If Node = "Recibos Egreso" Then
        repingre.acu = "V"
        repingre.Show 1

    End If

    If Node = "Caja Registradora Touch Screen" Then
        menucaja.Label3 = "CAJERO"
        menucaja.acu = "C"
        menucaja.Label1.Visible = True
        menucaja.turno.Visible = True
        menucaja.Label5 = "CAJA"
        menucaja.tipoterminal = "TOUCH"
        menucaja.Show 1

    End If

    If Node = "Caja Registradora Parqueo" Then
        menucaja.Label3 = "CAJERO"
        menucaja.acu = "C"
        menucaja.Label1.Visible = True
        menucaja.turno.Visible = True
        menucaja.Label5 = "CAJA"
        menucaja.tipoterminal = "PARKING"
        menucaja.Show 1

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

    If Node = "Salon" Then

        tsalon.Show 1

    End If

    If Node = "Mesa" Then
        tmesa.Show 1

    End If

    If Node = "Salon Numero/Mesas/Habit" Then
        tmesa.Show 1

    End If

End Sub

Function busca_clave1(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from vendedor where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_clave1 = Trim("" & mytablex.Fields("vevend"))

    End If

    mytablex.Close

End Function

