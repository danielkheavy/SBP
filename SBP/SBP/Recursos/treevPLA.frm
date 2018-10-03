VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form treevpla 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Planillas"
   ClientHeight    =   8580
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
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
      Height          =   825
      Left            =   600
      Picture         =   "treevPLA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir todo"
      Top             =   105
      Width           =   1695
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   7350
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   12965
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
            Picture         =   "treevPLA.frx":08CA
            Key             =   "picture1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "treevPLA.frx":0E64
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
Attribute VB_Name = "treevpla"
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
    treevpla.Hide
    Unload treevpla

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
    
    '''27/07/2017 kenyo Testing Completo al Sistema
    '    TreeView1.Nodes.Add sp, tvwChild, sh, "Concepto", "picture1"
    '    TreeView1.Nodes.Add sp, tvwChild, sh, "Modelos", "picture1"
    '    TreeView1.Nodes.Add sp, tvwChild, sh, "Periodo", "picture1"
    '''27/07/2017 kenyo Testing Completo al Sistema
    
    '' 03/07/2018 Conteo Fisico Sistema
    TreeView1.Nodes.Add sp, tvwChild, sh, "Periodo", "picture1"
    '' 03/07/2018 Conteo Fisico Sistema
 
    TreeView1.Nodes.Add sp, tvwChild, sh, "Personal", "picture1"
    
    TreeView1.Nodes.Add , , sp3, "Procesos", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Control Acceso", "picture1"
    
    '''27/07/2017 kenyo Testing Completo al Sistema
    '    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Ingreso Planilla", "picture1"
    '    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Generacion Planilla", "picture1"
    '    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Planilla", "picture1"
    '''27/07/2017 kenyo Testing Completo al Sistema
    
    TreeView1.Nodes.Add , , sp2, "Reportes", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Asistencia", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Lista de Personal", "picture1"
    'TreeView1.Nodes.Add sp2, tvwChild, sh2, "Asistencia Liquidacion"
    
    '''27/07/2017 kenyo Testing Completo al Sistema
    'TreeView1.Nodes.Add sp2, tvwChild, sh2, "Planilla ", "picture1"
    'TreeView1.Nodes.Add sp2, tvwChild, sh2, "Planilla ", "picture1"
    '''27/07/2017 kenyo Testing Completo al Sistema
    
    For I = 1 To 50
        buffer(I) = ""
    Next I
    
    '''27/07/2017 kenyo Testing Completo al Sistema
    'TreeView1.Nodes.Add , , sp4, "ReportesUsuario", "picture1"
    '''27/07/2017 kenyo Testing Completo al Sistema
    '------------------
    '    jindx = 0
    '    If mytablex.State = 1 Then mytablex.Close
    '   mytablex.Open "select * from archivo where menu='PLANILLA' and   estado='S'", cn, adOpenStatic, adLockOptimistic
    '   If mytablex.RecordCount > 0 Then
    '        Do
    '        If mytablex.EOF Then Exit Do
    '        jindx = jindx + 1
    '        buffer(jindx) = Trim("" & mytablex.Fields("descripcio"))
    '        TreeView1.Nodes.Add sp4, tvwChild, sh4, Trim("" & mytablex.Fields("descripcio")), "picture1"
    '        mytablex.MoveNext
    '        Loop
    '   End If
    '   mytablex.Close
    
    For I = 2 To TreeView1.Nodes.count - 1
        'TreeView1.Nodes(i).ExpandedImage = "Open"
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

    If Node = "Personal" Then
   
        tpersona.Show 1

    End If

    If Node = "Concepto" Then
        tplaconc.Show 1

    End If

    If Node = "Modelos" Then
        tmopla.Show 1

    End If

    If Node = "Periodo" Then
        tplape.Show 1

    End If

    If Node = "Control Acceso" Then
        tingper.Show 1

    End If

    If Node = "Ingreso Planilla" Then
        tplamo.Show 1

    End If

    If Node = "Generacion Planilla" Then

        opcion2 = "1"
        planilag.Show 1

    End If

    If Node = "Planilla" Then
        tplagepe.Show 1

    End If

    If Node = "Asistencia" Then
        opcion2 = "1"
        trepasis.Show 1

    End If

    If Node = "Asistencia Liquidacion" Then
        opcion2 = "2"
        trepasis.Show 1

    End If

    If Node = "Planilla " Then
        tplagepe.Show 1

    End If

    '''27/07/2017 kenyo Testing Completo al Sistema
    If Node = "Lista de Personal" Then
        treppersonal.Show 1

    End If

    '''27/07/2017 kenyo Testing Completo al Sistema

End Sub

