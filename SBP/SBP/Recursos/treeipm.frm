VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form treeipm 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importaciones"
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
      Picture         =   "treeipm.frx":0000
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
      Width           =   10215
      _ExtentX        =   18018
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
            Picture         =   "treeipm.frx":08CA
            Key             =   "picture1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "treeipm.frx":0E64
            Key             =   "picture2"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   8175
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF80&
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
Attribute VB_Name = "treeipm"
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
    treeipm.Hide
    Unload treeipm

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

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    sp = "sp"
    sp1 = "sp1"
    sp2 = "sp2"
    sp3 = "sp3"
    sp4 = "sp4"
    
    TreeView1.ImageList = ImageList1
    
    TreeView1.Nodes.Add , , sp, "Tablas", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Agencias de Aduana", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Gastos", "picture1"
    
    TreeView1.Nodes.Add , , sp3, "Procesos", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "FacturaGastos", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "FacturaDua", "picture1"
    
    TreeView1.Nodes.Add , , sp2, "Reportes", "picture1"
    'TreeView1.Nodes.Add sp2, tvwChild, sh2, "Facturacion Gastos ", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Dua ", "picture1"
    
    For I = 1 To 50
        buffer(I) = ""
    Next I
    
    For I = 1 To TreeView1.Nodes.count - 1
        TreeView1.Nodes(I).Expanded = True
    Next I
    
    '------------------
    If mytablex.State = 1 Then mytablex.Close
    jindx = 0
    mytablex.Open "select * from archivo where menu='CUENTAPAGAR' and   estado='S'", cn, adOpenStatic, adLockOptimistic

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

    If Node = "Agencias de Aduana" Then
        taduana.Show 1

    End If

    If Node = "Gastos" Then
        tADUANAG.Show 1

    End If

    If Node = "FacturaGastos" Then
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
        explorap.importacion = "GASTOS"
        explorap.Caption = "FACTURAGASTOS"

        explorap.Show 1

    End If

    If Node = "FacturaDua" Then
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
        explorap.importacion = "IMPORTACION"
        explorap.Caption = "FACTURADUA"

        explorap.Show 1

    End If

End Sub

