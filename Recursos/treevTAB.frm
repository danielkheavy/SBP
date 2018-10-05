VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form treevtab 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tablas"
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
      Height          =   705
      Left            =   0
      Picture         =   "treevTAB.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir todo"
      Top             =   0
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
            Picture         =   "treevTAB.frx":08CA
            Key             =   "picture1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "treevTAB.frx":0E64
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
      Height          =   960
      Left            =   -15
      TabIndex        =   1
      Top             =   15
      Width           =   12975
   End
   Begin VB.Menu d89 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "treevtab"
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
    treevtab.Hide
    Unload treevtab

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

    Dim sh4      As String

    Dim sp4      As String

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    sp = "sp"
    sp1 = "sp1"
    sp2 = "sp2"
    sp3 = "sp3"
    sp4 = "sp4"
    
    TreeView1.ImageList = ImageList1
    TreeView1.Nodes.Add , , sp, "Tablas", "picture1"
     
    TreeView1.Nodes.Add sp, tvwChild, sh, "Empresa", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Locales", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Almacenes", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Familias", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Secciones", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Categoria", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Marca", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Linea", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Color", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Grupo Inv.Fisico", "picture1"
    
    '''27/07/2017 kenyo Testing Completo al Sistema
    'TreeView1.Nodes.Add sp, tvwChild, sh, "Situaciones", "picture1"
    'TreeView1.Nodes.Add sp, tvwChild, sh, "Rentabilidad", "picture1"
    '''27/07/2017 kenyo Testing Completo al Sistema
    
    TreeView1.Nodes.Add sp, tvwChild, sh, "Ubicaciones", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Productos", "picture1"
     
    '''27/07/2017 kenyo Testing Completo al Sistema
    'TreeView1.Nodes.Add sp, tvwChild, sh, "Grupo Combinacion", "picture1"
    '''27/07/2017 kenyo Testing Completo al Sistema
        
    'TreeView1.Nodes.Add sp, tvwChild, sh, "Combinacion", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Personal", "picture1"
     
    '''27/07/2017 kenyo Testing Completo al Sistema
    'TreeView1.Nodes.Add sp, tvwChild, sh, "CodigoBarras", "picture1"
    '''27/07/2017 kenyo Testing Completo al Sistema
     
    TreeView1.Nodes.Add sp, tvwChild, sh, "Conceptos", "picture1"
    'TreeView1.Nodes.Add sp, tvwChild, sh, "Habitaciones", "picture1"
 
    '''27/07/2017 kenyo Testing Completo al Sistema
    'TreeView1.Nodes.Add sp, tvwChild, sh, "Parametros Reportes", "picture1"
    ' TreeView1.Nodes.Add sp, tvwChild, sh, "Ip Mquinas", "picture1"
    '''27/07/2017 kenyo Testing Completo al Sistema
 
    TreeView1.Nodes.Add sp, tvwChild, sh, "Tipos", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Clasificacion", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Clasificacion Sunat", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Clientes", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Proveedor", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Centro Costo", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Transportista", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Forma Pago", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Bancos", "picture1"
 
    '''27/07/2017 kenyo Testing Completo al Sistema
    'TreeView1.Nodes.Add sp, tvwChild, sh, "Secciones Cartera", "picture1"
    '''27/07/2017 kenyo Testing Completo al Sistema
 
    TreeView1.Nodes.Add sp, tvwChild, sh, "Zonas", "picture1"
 
    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
    TreeView1.Nodes.Add sp, tvwChild, sh, "Tipo Nota de Crédito/Débito", "picture1"
    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
  
    TreeView1.Nodes.Add sp, tvwChild, sh, "Tipo Cambio", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Tipo Documento", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Parametros Generales", "picture1"
     
    For I = 1 To 50
        buffer(I) = ""
    Next I
    
    ' 27/07/2017 kenyo Testing Completo al Sistema
    ' TreeView1.Nodes.Add , , sp4, "ReportesUsuario", "picture1", "picture1"
    ' 27/07/2017 kenyo Testing Completo al Sistema
    
    '------------------
    jindx = 0

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from archivo where menu='TABLAS' and   estado='S'", cn, adOpenStatic, adLockOptimistic

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
    
    '''27/07/2017 kenyo Testing Completo al Sistema
    For I = 1 To TreeView1.Nodes.count - 1
        'TreeView1.Nodes(i).ExpandedImage = "Open"
        TreeView1.Nodes(I).Expanded = True
    Next I

    Exit Sub
    '''27/07/2017 kenyo Testing Completo al Sistema
    
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

    If Node = "Clasificacion Sunat" Then
        tclsunat.Show 1

    End If

    'If Node = "Habitaciones" Then
    'thabita.Show 1
    'End If
    If Node = "Situaciones" Then
        tsituaci.Show 1

    End If

    If Node = "Empresa" Then
        tempresa.Show 1

    End If

    If Node = "Parametros Reportes" Then

        'pararep.Show 1
    End If

    If Node = "Locales" Then
        ttlocal.Show 1

    End If

    If Node = "Conceptos" Then
        tcptoie.Show 1

    End If

    If Node = "Grupo Inv.Fisico" Then
        tpefisic.Show 1

    End If

    If Node = "Almacenes" Then
        talmacen.Show 1

    End If

    'If Node = "Combinacion" Then
    'tabcombo.Show 1
    'End If
    If Node = "Familias" Then
        ttfamilia.Show 1

    End If

    If Node = "CodigoBarras" Then
        '''27/07/2017 kenyo Testing Completo al Sistema
        'tcobarra.Show 1
        '''27/07/2017 kenyo Testing Completo al Sistema

    End If

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    If Node = "Tipo Nota de Crédito/Débito" Then
        TIPONCD.Show 1

    End If

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
  
    If Node = "Secciones" Then
        tseccion.Show 1

    End If

    If Node = "Categoria" Then
        tcategor.Show 1

    End If

    If Node = "Marca" Then
        tnmarca.Show 1

    End If

    If Node = "Linea" Then
        tlinea.Show 1

    End If

    If Node = "Color" Then
        tncolor.Show 1

    End If

    If Node = "Ip Mquinas" Then
        tips.Show 1

    End If

    If Node = "Rentabilidad" Then
        tmargen.Show 1

    End If

    If Node = "Ubicaciones" Then
        tubica.Show 1

    End If

    If Node = "Productos" Then
        xprodet.Show 1

    End If

    If Node = "Personal" Then
        If busca_clave1(gusuario) <> "S" Then
            MsgBox "No tiene Permiso", 48, "Aviso"
            Exit Sub

        End If

        tpersona.Show 1

    End If

    If Node = "Tipos" Then
        txpoclie.Show 1

    End If

    If Node = "Clasificacion" Then
        tclasifi.Show 1

    End If

    If Node = "Clientes" Then
        tnclie.DBPROV = "clientes"
        tnclie.fdlo893.Visible = True
        tnclie.Show 1

    End If

    If Node = "Proveedor" Then
        tnclie.Label18.Enabled = True
        tnclie.Caption = "Tabla de proveedores"
        tnclie.DBPROV = "Proveedo"
        tnclie.Show 1

    End If

    If Node = "Centro Costo" Then
        tccosto.Show 1

    End If

    If Node = "Transportista" Then
        ttranspo.Show 1

    End If

    If Node = "Forma Pago" Then
        tfpago.Show 1

    End If

    If Node = "Bancos" Then
        tbanco.Show 1

    End If

    If Node = "Secciones Cartera" Then
        tcase.Show 1

    End If

    If Node = "Zonas" Then
        tnzona.Show 1

    End If

    If Node = "Tipo Cambio" Then
        tcambio.Show 1

    End If

    If Node = "Tipo Documento" Then
        tdocumen.Show 1

    End If

    If Node = "Parametros Generales" Then
        ttparame.Show 1

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

