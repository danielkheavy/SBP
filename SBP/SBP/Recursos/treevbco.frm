VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form treevbco 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bancos"
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
      Left            =   8640
      Picture         =   "treevbco.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir todo"
      Top             =   120
      Width           =   1695
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5640
      Top             =   2880
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
            Picture         =   "treevbco.frx":08CA
            Key             =   "picture1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "treevbco.frx":0E64
            Key             =   "picture2"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   8175
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   14420
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Scroll          =   0   'False
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
Attribute VB_Name = "treevbco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnsalir_Click()
    d89_Click

End Sub

Private Sub d89_Click()
    treevbco.Hide
    Unload treevbco

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

    Dim I   As Integer

    sp = "sp"
    sp1 = "sp1"
    sp2 = "sp2"
    sp3 = "sp3"
    TreeView1.ImageList = ImageList1
   
    TreeView1.Nodes.Add , , sp, "Tablas", "picture1"
    
    TreeView1.Nodes.Add sp, tvwChild, sh, "Bancos", "picture2"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Clientes", "picture2"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Proveedores", "picture2"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Personal", "picture2"
    
    TreeView1.Nodes.Add , , sp3, "Procesos", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Bancos ", "picture2"
    
    For I = 2 To TreeView1.Nodes.count - 1
        TreeView1.Nodes(I).Expanded = True
    Next I
    
    Exit Sub

    'cmdLlenarTree_Click

End Sub
 
Private Sub Image1_Click()

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

    If Node = "Clientes" Then
        tnclie.DBPROV = "Clientes"
        tnclie.Show 1

    End If

    If Node = "Proveedores" Then
        tnclie.DBPROV = "Proveedo"
        tnclie.Show 1

    End If

    If Node = "Bancos" Then
   
        tbanco.Show 1

    End If

    If Node = "Bancos " Then
        tmovcheq.Show 1

    End If

End Sub
