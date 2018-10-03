VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form treevcon 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contabilidad"
   ClientHeight    =   9240
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   14625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   14625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Mes de Trabajo"
      Height          =   3735
      Left            =   600
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   9255
      Begin VB.TextBox ANOCONTA 
         Height          =   615
         Left            =   3600
         MaxLength       =   4
         TabIndex        =   10
         Top             =   2760
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         Height          =   495
         Left            =   7800
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Graba"
         Height          =   495
         Left            =   7800
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox mesconta 
         Height          =   615
         Left            =   3600
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AÑO DE TRABAJO"
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Top             =   2280
         Width           =   3135
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MES DE TRABAJO"
         Height          =   255
         Left            =   3240
         TabIndex        =   8
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00 01 02 03 04 05 06 07 08 09 10 11 12 13"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   7
         Top             =   840
         Width           =   4455
      End
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
      Picture         =   "treevcon.frx":0000
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
            Picture         =   "treevcon.frx":08CA
            Key             =   "picture1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "treevcon.frx":0E64
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
      Width           =   12975
   End
   Begin VB.Menu d89 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "treevcon"
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

Private Sub Command1_Click()

    Dim mytablex As New ADODB.Recordset

    mesconta = Format(Val(mesconta), "00")

    If Val(mesconta) < 0 And Val(mesconta) > 13 Then
        mesconta = ""
        Exit Sub

    End If

    If Not IsNumeric(ANOCONTA) Then
        ANOCONTA = ""
        Exit Sub

    End If

    If Len(ANOCONTA) < 4 Then
        ANOCONTA = ""
        Exit Sub

    End If

    mytablex.Open "select * from parame where codigo='01' ", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        mytablex.Fields("mesconta") = mesconta
        mytablex.Fields("anoconta") = ANOCONTA
        mytablex.Update

    End If

    mytablex.Close
    Frame1.Visible = False

End Sub

Private Sub Command2_Click()
    Frame1.Visible = False

End Sub

Private Sub d89_Click()

    If Frame1.Visible = True Then
        Frame1.Visible = False
        Exit Sub

    End If

    treevcon.Hide
    Unload treevcon

End Sub

Private Sub Form_Load()

    'ojo
    'ventas
    '   debe 12123           Boleta  Factura 12121
    ' igv 40111
    ' Subtotal 70211

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

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    sp = "sp"
    sp1 = "sp1"
    sp2 = "sp2"
    sp3 = "sp3"
    sp4 = "sp4"
    
    TreeView1.ImageList = ImageList1
    
    TreeView1.Nodes.Add , , sp, "Tablas", "picture1"
    
    'TreeView1.Nodes.Add sp, tvwChild, sh, "Plan de Cuentas Predefinido", "picture1"
    'TreeView1.Nodes.Add sp, tvwChild, sh, "Proveedores", "picture1"
    'TreeView1.Nodes.Add sp, tvwChild, sh, "Clientes", "picture1"
    'TreeView1.Nodes.Add sp, tvwChild, sh, "Personal", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Tipos Documentos", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Centro de Costo", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Origenes", "picture1"
    'TreeView1.Nodes.Add sp, tvwChild, sh, "Libro Auxiliar", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Tipo De Cambio Diario", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Tipos de Cuentas", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Plan de Cuentas", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Cambio Mes Trabajo", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Parametros de Cuentas", "picture1"
       
    TreeView1.Nodes.Add , , sp3, "Procesos", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Voucher", "picture1"
    'TreeView1.Nodes.Add sp3, tvwChild, sh3, "Plan de Cuentas", "picture1"
    'TreeView1.Nodes.Add sp3, tvwChild, sh3, "Asiento Contable", "picture1"
    'TreeView1.Nodes.Add sp3, tvwChild, sh3, "Asientos Prefinidos", "picture1"
    
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "InterfaseSiscont", "picture1"
    
    TreeView1.Nodes.Add , , sp2, "Reportes", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Reporte Diario", "picture1"
    'TreeView1.Nodes.Add sp2, tvwChild, sh2, "Libro Diario", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Libro Diario Sunat", "picture1"
    'TreeView1.Nodes.Add sp2, tvwChild, sh2, "Libro Diario x Fuente", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Libro Mayor", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Balance Prueba", "picture1"
    'TreeView1.Nodes.Add sp2, tvwChild, sh2, "Balance Comprobacion", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Resumen de Gastos", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Ganancias y Perdidas", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Balance tributario", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Auxiliares por Ruc", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Auxliares por Fuente", "picture1"
    'TreeView1.Nodes.Add sp3, tvwChild, sh3, "Interfase Orion", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Interfase Siscont Exportar", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "JalarDesdeSiscont", "picture1"
    
    For I = 1 To 50
        buffer(I) = ""
    Next I

    TreeView1.Nodes.Add , , sp4, "ReportesUsuario", "picture1"
    
    '------------------
    If mytablex.State = 1 Then mytablex.Close
    jindx = 0
    mytablex.Open "select * from archivo where menu='CONTABILIDAD' and   estado='S'", cn, adOpenStatic, adLockOptimistic

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
    
    For I = 1 To TreeView1.Nodes.count - 1
        'TreeView1.Nodes(i).ExpandedImage = "Open"
        TreeView1.Nodes(I).Expanded = True
    Next I
    
    Exit Sub

    'cmdLlenarTree_Click

End Sub
 
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

    Dim mytablex As New ADODB.Recordset

    If Node = "Cambio Mes Trabajo" Then
        Frame1.Visible = True
        mesconta = ""
        ANOCONTA = ""
        mytablex.Open "select * from parame where codigo='01' ", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            mesconta = Trim("" & mytablex.Fields("mesconta"))
            ANOCONTA = Trim("" & mytablex.Fields("ANOconta"))

        End If

        mytablex.Close
        Exit Sub
   
    End If

    If Node = "Tipo De Cambio Diario" Then
        tcambio.Show 1

    End If

    If Node = "Parametros de Cuentas" Then
        tcontapa.Show 1

    End If

    If Node = "Proveedores" Then
        tnclie.DBPROV = "Proveedo"
        tnclie.Show 1

    End If

    If Node = "Clientes" Then
        tnclie.DBPROV = "Clientes"
        tnclie.Show 1

    End If

    If Node = "Personal" Then
        tpersona.Show 1

    End If

    If Node = "Voucher" Then
        tsiscont.Show 1

    End If

    If Node = "JalarDesdeSiscont" Then
        txpsisco.Show 1

    End If

    If Node = "InterfaseSiscont" Then
        tsisint.Show 1

    End If

    'If Node = "Asientos Prefinidos" Then
    'tpreasto.Show 1
    'End If

    'If Node = "InicioContable" Then
    '   tparacon.Show 1
    '
    'End If
    If Node = "Centro de Costo" Then
        tccosto.Show 1

    End If

    If Node = "Plan de Cuentas" Then
        tctable.Show 1

    End If

    'If Node = "Asiento Contable" Then

    'tasiento.Fechai = "01/01/" & Format(Year(Now), "0000")
    'tasiento.fechaf = "30/12/" & Format(Year(Now), "0000")

    'tasiento.Show 1
    'End If
    If Node = "Tipos de Cuentas" Then
        ttipocta.Show 1

    End If

    'If Node = "Libro Diario" Then
    'tlbrodia.tipo = "CUENTA"
    'tlbrodia.tiporeporte = "NORMAL"
    'tlbrodia.Show 1
    'End If
    'If Node = "Reporte Diario" Then
    'tlbrodia.tipo = "CUENTA"
    'tlbrodia.tiporeporte = "NORMAL"
    'tlbrodia.Show 1
    'End If
    'If Node = "Libro Diario Sunat" Then
    'tlbrodia.tipo = "CUENTA"
    'tlbrodia.digitos = "%"
    'tlbrodia.tiporeporte = "SUNAT"
    'tlbrodia.Show 1
    'End If
    'If Node = "Libro Mayor" Then
    'tlbrodia.digitos = "15"
    'tlbrodia.tipo = "CUENTA"
    'tlbrodia.Caption = "LIBRO MAYOR"
    'tlbrodia.tiporeporte = "MAYOR"
    'tlbrodia.Show 1
    'End If
    'If Node = "Balance Prueba" Then
    'tlbrodia.digitos = "2"
    'tlbrodia.tipo = "CUENTA"
    'tlbrodia.Caption = "BALANCE PRUEBA"
    'tlbrodia.tiporeporte = "BALANCEPRUEBA"
    'tlbrodia.Show 1
    'End If
    'If Node = "Centralizacion" Then
    'genconta.Show 1
    'End If
    If Node = "Origenes" Then
        tfuentec.Show 1

    End If

    'If Node = "Plan de Cuentas Predefinido" Then
    'tctadef.Show 1
    'End If
    'If Node = "Cuentas de Enlace" Then
    '   tenlace.Show 1
    'End If
    If Node = "Tipos Documentos" Then
        tdocta.Show 1

    End If

End Sub

