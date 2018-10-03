VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form treevc 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compras"
   ClientHeight    =   9240
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   11670
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
      Picture         =   "TREEVC.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir todo"
      Top             =   120
      Width           =   1695
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   14420
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
      Left            =   7680
      Top             =   360
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
            Picture         =   "TREEVC.frx":08CA
            Key             =   "picture1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TREEVC.frx":0E64
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
      TabIndex        =   1
      Top             =   0
      Width           =   11535
   End
   Begin VB.Menu d89 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "treevc"
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
    treevc.Hide
    Unload treevc

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

    Dim mytablex As New ADODB.Recordset

    Dim I        As Integer

    sp = "sp"
    sp1 = "sp1"
    sp2 = "sp2"
    sp3 = "sp3"
    sp4 = "sp4"
    sp5 = "sp5"
    
    TreeView1.ImageList = ImageList1
    
    TreeView1.Nodes.Add , , sp, "Tablas", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Locales", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Almacenes", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Tipo Documentos", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Proveedor", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Transportistas", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Productos", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Forma Pago", "picture1"
    
    TreeView1.Nodes.Add , , sp1, "Movimientos", "picture1"
    TreeView1.Nodes.Add sp1, tvwChild, sh1, "Guia Remision", "picture1"
    TreeView1.Nodes.Add sp1, tvwChild, sh1, "Factura Compras", "picture1"
    TreeView1.Nodes.Add sp1, tvwChild, sh1, "Nota Credito", "picture1"
    TreeView1.Nodes.Add sp1, tvwChild, sh1, "Nota Debito", "picture1"
    
    TreeView1.Nodes.Add , , sp3, "Procesos", "picture1"
    
    '''27/07/2017 kenyo Testing Completo al Sistema
    'TreeView1.Nodes.Add sp3, tvwChild, sh3, "Generacion Orden Compra", "picture1"
    '''27/07/2017 kenyo Testing Completo al Sistema
    
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Orden Compra", "picture1"
    
    TreeView1.Nodes.Add , , sp2, "Reportes", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Orden Compra ", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Guia Remision ", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Factura Compra ", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Registro Compras", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Factura Compra Productos ", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Lista Proveedores ", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Margenes", "picture1"
    'TreeView1.Nodes.Add sp2, tvwChild, sh2, "Cuentas por Pagar", "picture1"
    'TreeView1.Nodes.Add sp2, tvwChild, sh2, "Letras por Pagar", "picture1"
    
    TreeView1.Nodes.Add , , sp4, "Estadisticas", "picture1"
    TreeView1.Nodes.Add sp4, tvwChild, sh4, "Compras", "picture1"
    TreeView1.Nodes.Add sp4, tvwChild, sh4, "Compras Mensuales", "picture1"
    TreeView1.Nodes.Add sp4, tvwChild, sh4, "Ranking", "picture1"
    TreeView1.Nodes.Add sp4, tvwChild, sh4, "Graficos", "picture1"
     
    For I = 1 To 50
        buffer(I) = ""
    Next I
    
    '''27/07/2017 kenyo Testing Completo al Sistema
    '
    '    TreeView1.Nodes.Add , , sp5, "ReportesUsuario", "picture1"
    '
    '    '------------------
    '    jindx = 0
    '    If mytablex.State = 1 Then mytablex.Close
    '   mytablex.Open "select * from archivo where menu='COMPRAS' and   estado='S'", cn, adOpenStatic, adLockOptimistic
    '   If mytablex.RecordCount > 0 Then
    '        Do
    '        If mytablex.EOF Then Exit Do
    '        jindx = jindx + 1
    '        buffer(jindx) = Trim("" & mytablex.Fields("descripcio"))
    '        TreeView1.Nodes.Add sp5, tvwChild, sh5, Trim("" & mytablex.Fields("descripcio")), "picture1"
    '        mytablex.MoveNext
    '        Loop
    '   End If
    '   mytablex.Close
    '
    '     For i = 2 To TreeView1.Nodes.count - 1
    'TreeView1.Nodes(i).Expanded = True
    'Next i
    '''27/07/2017 kenyo Testing Completo al Sistema

    Exit Sub

    'cmdLlenarTree_Click

End Sub
 
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

    Dim I As Integer

    'If jindx > 0 Then
    'For i = 1 To jindx
    '    If Len(buffer(i)) > 0 Then
    '       If Node = buffer(i) Then
    '          ejecuta_reporte buffer(i)
    '       End If
    '    End If
    'Next i
    'End If
    If Node = "Registro Compras" Then
        'treporte.tituloreporte = "Registro de Compras"
        'treporte.archivoreporte = globaldir & "\reportes\registroventas.rpt"
        'treporte.acu = "C"
        'treporte.Show 1

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

    If Node = "Factura Compras" Then
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
        explorap.fechaf = Format(Now, "dd/mm/yyyy")
        explorap.fechai = Format(Now, "dd/mm/yyyy")
        explorap.Caption = "Documentos Facturacion Compras"
        explorap.tipoclie = "P"
        explorap.acu = "C"
        explorap.importacion = "COMERCIAL"
        explorap.Show 1

    End If

    If Node = "Guia Remision" Then
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
        explorap.fechaf = Format(Now, "dd/mm/yyyy")
        explorap.fechai = Format(Now, "dd/mm/yyyy")

        explorap.Caption = "Documentos Guia Remision Compra"
        explorap.tipoclie = "P"
        explorap.acu = "S"
        explorap.Show 1

    End If

    If Node = "Nota Credito" Then
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
        explorap.Caption = "Documentos Nota Credito Compras"
        explorap.tipoclie = "P"
        explorap.acu = "N"
        explorap.Show 1

    End If

    If Node = "Nota Debito" Then
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        'inicio 10/02/2018
        'explorap.fk4844.Visible = False
        'fin 10/02/2018

        explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
        explorap.fechaf = Format(Now, "dd/mm/yyyy")
        explorap.fechai = Format(Now, "dd/mm/yyyy")
        explorap.Caption = "Documentos Nota Debito Compras"
        explorap.tipoclie = "P"
        explorap.acu = "O"
        explorap.Show 1

    End If

    If Node = "Generacion Orden Compra" Then

        Dim found As Integer

        found = copiar_temporalpe()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
            End
            Exit Sub

        End If

        dgusuario = "_t" & gusuario
        tpedauto.Show 1

    End If

    If Node = "Orden Compra" Then
        cgusuario = "cordenc"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "dordenc"
        explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
        explorap.fechaf = Format(Now, "dd/mm/yyyy")
        explorap.Caption = "Documentos Orden de Compra"
        explorap.tipoclie = "P"
        explorap.acu = "R"
        explorap.Show 1

    End If

    If Node = "Orden Compra " Then
        cgusuario = "CORDENC"
        dgusuariog = "DORDENC"
        repdocum.acu = "R"
        repdocum.Show 1

    End If

    If Node = "Guia Remision " Then
        cgusuario = "FACTURA"
        dgusuariog = "DETALLE"
        repdocum.acu = "S"
        repdocum.Show 1

    End If

    If Node = "Factura Compra " Then
        cgusuario = "FACTURA"
        dgusuariog = "DETALLE"
        repdocum.acu = "C"
        repdocum.Show 1

    End If

    If Node = "Factura Compra Productos " Then
        opcion2 = "1"
        repraped.acu = "C" 'PEDIDO
        repraped.xdata = "DETALLE"
        repraped.Show 1
        'treporte.tituloreporte = "Reporte Compras Detalle"
        'treporte.archivoreporte = globaldir & "\reportes\facturadetalle.rpt"
        'treporte.acu = "C"
        'treporte.Show 1

    End If

    If Node = "Registro Compras" Then

        cgusuario = "FACTURA"
        dgusuariog = "DETALLE"
        repdocrv.titulo = "REGISTRO DE COMPRAS " & dicmoneda
        repdocrv.acu = "C"
        repdocrv.Show 1

    End If

    If Node = "Lista Proveedores " Then
        trepprov.Show 1

    End If

    If Node = "Margenes" Then
        opcion2 = "94"
        repinv.excell.Visible = True

        '''27/07/2017 kenyo Testing Completo al Sistema
        repinv.excell.Value = False
        '''27/07/2017 kenyo Testing Completo al Sistema

        repinv.Label17.Visible = True
        repinv.Combo1.Visible = True
        repinv.Label25.Visible = True
        repinv.gcanti.Visible = True
        repinv.Show 1

    End If

    If Node = "Compras" Then
        opcion2 = "11"   'analisis de ventas
        cgusuario = "FACTURA"
        dgusuariog = "DETALLE"
        'repdocum.Label18.Visible = False
        'repdocum.Combo1.Visible = False
        repdocum.vdetalle.Enabled = False
        repdocum.vfpago.Enabled = False
        repdocum.acu = "C"
        repdocum.Show 1

    End If

    If Node = "Compras Mensuales" Then
        opcion2 = "12"   'analisis de ventas
        cgusuario = "FACTURA"
        dgusuariog = "DETALLE"
        'repdocum.Label18.Visible = False
        'repdocum.Combo1.Visible = False
        repdocum.vdetalle.Enabled = False
        repdocum.vfpago.Enabled = False
        repdocum.acu = "C"
        repdocum.Show 1

    End If

    If Node = "Ranking" Then
        opcion2 = "2"
        repraped.Label12.Visible = True
        repraped.orden.Visible = True
        repraped.acu = "C" 'PEDIDO
        repraped.xdata = "DETALLE"
        repraped.Show 1

    End If

    If Node = "Graficos" Then
        FrmChart.acu = "C"
        FrmChart.Show 1

    End If

    If Node = "Cuentas por Pagar" Then
        repctaxc.acu = "C"
        repctaxc.Show 1

    End If

    If Node = "Letras por Pagar" Then
        REPLETRA.titulo = "Letras por Pagar"
        REPLETRA.acu = "C"
        REPLETRA.Show 1

    End If

End Sub

