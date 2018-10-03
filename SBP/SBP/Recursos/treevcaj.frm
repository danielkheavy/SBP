VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form treevcaj 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Caja"
   ClientHeight    =   9240
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   10620
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
      Left            =   0
      Picture         =   "treevcaj.frx":0000
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
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   14420
      _Version        =   393217
      LineStyle       =   1
      PathSeparator   =   "*"
      Style           =   7
      BorderStyle     =   1
      Appearance      =   0
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
            Picture         =   "treevcaj.frx":08CA
            Key             =   "picture1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "treevcaj.frx":0E64
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
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   135
   End
   Begin VB.Menu d89 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "treevcaj"
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
    treevcaj.Hide
    Unload treevcaj

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
    
    TreeView1.ImageList = ImageList1
    
    TreeView1.Nodes.Add , , sp, "Tablas", "picture1"
    
    TreeView1.Nodes.Add sp, tvwChild, sh, "Proveedores", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Clientes", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Personal", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Forma Pago", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Concepto", "picture1"
    
    TreeView1.Nodes.Add , , sp3, "Procesos", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Ingreso Dinero", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Egreso Dinero", "picture1"
    
    TreeView1.Nodes.Add , , sp2, "Reportes", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Ingreso Dinero ", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Egreso Dinero ", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Movimiento Ingreso Egreso ", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Flujo de Dinero ", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Forma Pago ", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Cuadre Caja ", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Compras Ventas", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Flujo Entradas salidas", "picture1"
    
    For I = 1 To 50
        buffer(I) = ""
    Next I

    '''27/07/2017 kenyo Testing Completo al Sistema
    'TreeView1.Nodes.Add , , sp4, "ReportesUsuario", "picture1"
    '''27/07/2017 kenyo Testing Completo al Sistema
    '------------------
    jindx = 0

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from archivo where menu='CAJA' and   estado='S'", cn, adOpenStatic, adLockOptimistic

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
        TreeView1.Nodes(I).Expanded = True
    Next I
    
    Exit Sub

    'cmdLlenarTree_Click

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

    If Node = "Flujo de Dinero " Then
        tflfp.Show 1

    End If

    If Node = "Proveedores" Then
        tnclie.DBPROV = "Proveedo"
        tnclie.Show 1

    End If

    If Node = "Clientes" Then
        tnclie.DBPROV = "clientes"
        tnclie.Show 1

    End If

    If Node = "Personal" Then
   
        tpersona.Show 1

    End If

    If Node = "Concepto" Then
        tcptoie.Show 1

    End If

    If Node = "Forma Pago" Then
   
        tfpago.Show 1

    End If

    If Node = "Ingreso Dinero" Then
        explreci.xcuentaco = "cuentac"
        explreci.XCUENTACO1 = "cuentacd"
        explreci.Caption = "INGRESO DINERO"
        'explreci.afecta = "C"
        explreci.acu = "W"
        explreci.Show 1

    End If

    If Node = "Egreso Dinero" Then
        explreci.xcuentaco = "cuentap"
        explreci.XCUENTACO1 = "cuentapd"
        explreci.Caption = "EGRESO DINERO"
        'explreci.afecta = "P"  'proveedor
        explreci.acu = "V"
        explreci.Show 1

    End If

    If Node = "Ingreso Dinero " Then
        repingre.xcuentaco = "cuentac"
        repingre.XCUENTACO1 = "cuentacd"
        repingre.acu = "W"
        repingre.Show 1

    End If

    If Node = "Egreso Dinero " Then
        repingre.xcuentaco = "cuentap"
        repingre.XCUENTACO1 = "cuentapd"
        repingre.acu = "V"
        repingre.Show 1

    End If

    If Node = "Movimiento Ingreso Egreso " Then
        trecitot.Caption = "COMPROBANTES INGRESOS EGRESOS"
        'explreci.acu = "W"
        trecitot.Show 1

    End If

    If Node = "Forma Pago " Then
        repfpago.Show 1

    End If

    If Node = "Cuadre Caja " Then

        '''27/07/2017 kenyo Testing Completo al Sistema
        'flag_clave1 = 0
        'tconcla.X = "CUADRE"
        'tconcla.Show 1
        'If flag_clave1 <> 1 Then  '
        'Exit Sub
        'End If
        'opcion2 = "1"
        'opcion1 = "1"
        'opcion3 = "2"
        'usuariopos = gusuario
   
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

        '''27/07/2017 kenyo Testing Completo al Sistema

    End If

    If Node = "Compras Ventas" Then
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tcomvta.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
        tcomvta.fechaf = Format(Now, "dd/mm/yyyy")
        tcomvta.fechai = Format(Now, "dd/mm/yyyy")
        tcomvta.Caption = "Documentos Facturacion Compras Ventas"
        tcomvta.tipoclie = "%"
        tcomvta.acu = "%"
        tcomvta.Show 1

    End If

    If Node = "Flujo Entradas salidas" Then
        tflujoe.Show 1

    End If

End Sub

