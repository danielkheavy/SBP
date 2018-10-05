VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form treevpro 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Produccion"
   ClientHeight    =   9240
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   12300
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
      Left            =   8640
      Picture         =   "treevpro.frx":0000
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
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   14420
      _Version        =   393217
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
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
            Picture         =   "treevpro.frx":08CA
            Key             =   "picture1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "treevpro.frx":0E64
            Key             =   "picture2"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   5160
      TabIndex        =   3
      Top             =   1200
      Width           =   6855
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   8175
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   7095
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
Attribute VB_Name = "treevpro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnsalir_Click()
    d89_Click

End Sub

Private Sub d89_Click()
    treevpro.Hide
    Unload treevpro

End Sub

Private Sub Form_Load()

    Dim sp  As String

    Dim sh  As String

    Dim sp1 As String

    Dim sh1 As String

    Dim sp2 As String

    Dim sh2 As String

    Dim sp3 As String

    Dim sh3 As String

    Dim sp4 As String

    Dim sh4 As String
    
    sp = "sp"
    sp1 = "sp1"
    sp2 = "sp2"
    sp3 = "sp3"
    sp4 = "sp4"
    
    comentario
    
    TreeView1.ImageList = ImageList1
    TreeView1.Nodes.Add , , sp1, "Tablas", "picture1"
    
    'TreeView1.Nodes.Add sp1, tvwChild, sh1, "Estado"
    'TreeView1.Nodes.Add sp1, tvwChild, sh1, "Lote"
    'TreeView1.Nodes.Add sp1, tvwChild, sh1, "CentroCosto"
    TreeView1.Nodes.Add sp1, tvwChild, sh1, "TablasProduccion", "picture1"
    'TreeView1.Nodes.Add sp1, tvwChild, sh1, "AreasProduccion", "picture1"
    TreeView1.Nodes.Add sp1, tvwChild, sh1, "CentroProduccion", "picture1"
    
    TreeView1.Nodes.Add , , sp, "Planeamiento", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Productos", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Formulas", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Pedidos", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "OrdenTrabajo", "picture1"
    'TreeView1.Nodes.Add sp, tvwChild, sh, "Planificacion"
    
    TreeView1.Nodes.Add , , sp3, "Gesion Recursos", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "NotaIngreso", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "NotaSalida", "picture1"
    
    TreeView1.Nodes.Add , , sp2, "Ejecucion", "picture1"
    'TreeView1.Nodes.Add sp2, tvwChild, sh2, "RegistroTareo", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "ParteProduccion", "picture1"
    'TreeView1.Nodes.Add sp2, tvwChild, sh2, "NotificacionProceso", "picture1"
    
    TreeView1.Nodes.Add , , sp4, "Reportes", "picture1"
    TreeView1.Nodes.Add sp4, tvwChild, sh4, "Hoja de Costo", "picture1"
    TreeView1.Nodes.Add sp4, tvwChild, sh4, "Analisis Hoja de Costo", "picture1"
    TreeView1.Nodes.Add sp4, tvwChild, sh4, "Costos de Produccion", "picture1"
    
    For I = 1 To TreeView1.Nodes.count - 1
        TreeView1.Nodes(I).Expanded = True
    Next I
    
    Exit Sub
    
    TreeView1.Nodes.Add , , sp, "Tablas"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Secciones"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Operaciones"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Personal"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Componentes"
    
    TreeView1.Nodes.Add , , sp3, "Procesos"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Orden Produccion"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Tarjeta de Control"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Parte Produccion"
    
    TreeView1.Nodes.Add , , sp2, "Reportes"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Produccion Por Centro Produccion"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Produccion Por Maquina"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Eficiencia por Centro Produccion"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Eficiencia por Maquina"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Gastos de Fabricacion"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Registro Produccion"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Liquidacion Componentes/materiales"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Avance Produccion"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Rentabilidad por Pedido Venta"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Reportes Variables"
    
    Exit Sub

    'cmdLlenarTree_Click

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

    If Node = "CentroProduccion" Then
        ttcentro.Show 1

    End If

    If Node = "Pedidos" Then
        cgusuario = "CPEDIDOV"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DPEDIDOV"
        explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
        explorap.fechaf = Format(Now, "dd/mm/yyyy")
        explorap.Caption = "Documentos Pedidos Produccion"
        explorap.tipoclie = "C"
        explorap.acu = "I"
        explorap.Show 1

    End If

    If Node = "TablasProduccion" Then
        tablapro.Show 1

    End If

    If Node = "Productos" Then
        xprodet.Show 1

    End If

    If Node = "Formulas" Then
        ttrecepr.Show 1

    End If

    If Node = "CentroCosto" Then
        tccosto.Show 1

    End If

    If Node = "AreasProduccion" Then
        tarea.Show 1

    End If

    If Node = "OrdenTrabajo" Then
        ttordent.Show 1

    End If

    If Node = "Planificacion" Then
        tplanif.Show 1

    End If

    If Node = "NotaIngreso" Then
        'tnotaing.tipomov = "S"
        'tnotaing.Show 1
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

        explorap.Caption = "Documentos Guia Remision Compra"
        explorap.tipoclie = "V"
        explorap.acu = "S"
        explorap.tinterno = "S"
        explorap.Show 1
   
    End If

    If Node = "NotaSalida" Then
        'tnotaing.tipomov = "T"
        'tnotaing.Show 1
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        'inicio 10/02/2018`pll
        'explorap.fk4844.Visible = False
        'fin 10/02/2018`pll

        explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
        explorap.fechaf = Format(Now, "dd/mm/yyyy")
        explorap.fechai = Format(Now, "dd/mm/yyyy")
        explorap.Caption = "Documentos Guia Remision Ventas"
        explorap.tipoclie = "V"
        explorap.tinterno = "S"
        explorap.acu = "T"
        explorap.Show 1

    End If

    If Node = "ParteProduccion" Then
        ttordent.viene = "ParteProduccion"
        ttordent.Show 1

        'tpartepc.Show 1
    End If

    Exit Sub

    If Node = "Secciones" Then
        tprosecc.Show 1

    End If

    If Node = "Operaciones" Then
        toperaco.Show 1

    End If

    If Node = "Personal" Then
        tpersona.Show 1

    End If

    If Node = "Componentes" Then
        xprodet.Show 1

    End If

    If Node = "Tarjeta de Control" Then
        ttarprod.Show 1

    End If

    If Node = "Orden Produccion" Then
        tplanopr.Show 1

    End If

    If Node = "Parte Produccion" Then
        tsecpro.Show 1

        'tpartes.Show 1
    End If

    If Node = "Produccion" Then
        repprodu.titulo = "Reportes de Produccion"
        repprodu.Show 1

    End If

End Sub

Sub comentario()
    Label1 = "Operaciones de un proceso de Produccion" + Chr$(10) + Chr$(13)
    Label1 = Label1 & "Registro de Tablas para modulo de Produccion" + Chr$(10) + Chr$(13)
    Label1 = Label1 & "Registro Centro de Produccion" + Chr$(10) + Chr$(13)
    Label1 = Label1 & "Registro Formulacion o recetario" + Chr$(10) + Chr$(13)
    Label1 = Label1 & "1.Registro Ordenes Trabajo" + Chr$(10) + Chr$(13)
    Label1 = Label1 & "2.Retirar materiales para la Ot" + Chr$(10) + Chr$(13)
    Label1 = Label1 & "3.Registrar tareo de la mano de obra" + Chr$(10) + Chr$(13)
    Label1 = Label1 & "4.registrar parte de produccion" + Chr$(10) + Chr$(13)
    Label1 = Label1 & "5.revisar Hoja de Costo"

End Sub
