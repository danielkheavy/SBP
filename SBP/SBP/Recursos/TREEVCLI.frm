VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form treevcli 
   BackColor       =   &H80000013&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clinicas"
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
      Picture         =   "TREEVCLI.frx":0000
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
Attribute VB_Name = "treevcli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnsalir_Click()
    d89_Click

End Sub

Private Sub d89_Click()
    treevcli.Hide
    Unload treevcli

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

    sp = "sp"
    sp1 = "sp1"
    sp2 = "sp2"
    sp3 = "sp3"

    With TreeView1
        .Style = tvwTreelinesPlusMinusText
        .LineStyle = tvwRootLines
        .PathSeparator = "\"
        .Indentation = Screen.TwipsPerPixelX * 5 '256
        '
        ' No permitir la edición automática del texto
        .LabelEdit = tvwManual
        ' Para que se pueda expandir al seleccionar un nodo,
        ' cambia este valor a True,
        ' si se deja en False, tendrás que hacer doble-click
        .SingleSel = False
        ' Para que al perder el foco,
        ' se siga viendo el que está seleccionado
        .HideSelection = False
        '
        .refresh

    End With

    TreeView1.Nodes.Add , , sp, "Tablas"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Clientes"
    
    TreeView1.Nodes.Add , , sp1, "Procesos"
    TreeView1.Nodes.Add sp1, tvwChild, sh1, "Consulta"
    TreeView1.Nodes.Add sp1, tvwChild, sh1, "Diagnostico"
    TreeView1.Nodes.Add sp1, tvwChild, sh1, "Tratamiento"
    TreeView1.Nodes.Add sp1, tvwChild, sh1, "Asistencia"
    
    Exit Sub

    'cmdLlenarTree_Click

End Sub
 
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

    If Node = "Clientes" Then
        tnclie.DBPROV = "clientes"
        tnclie.Show 1

    End If

    If Node = "Consulta" Then
        tconsult.Show 1

    End If

    If Node = "Diagnostico" Then
        tdiagnos.Show 1

    End If

    If Node = "Tratamiento" Then
        ttratame.Show 1

    End If

    If Node = "Asistencia" Then
        tasiste.Show 1

    End If

End Sub
