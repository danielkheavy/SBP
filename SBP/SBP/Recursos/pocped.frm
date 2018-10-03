VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form pocped 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedidos"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10890
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7560
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   3600
      TabIndex        =   5
      Top             =   2640
      Width           =   3255
      Begin VB.TextBox clave 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   840
         MaxLength       =   2
         TabIndex        =   0
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "Entrar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Terminal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ingrese Numero Terminal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.TextBox producto 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   720
      MaxLength       =   5
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "pocped.frx":0000
      Height          =   2775
      Left            =   0
      OleObjectBlob   =   "pocped.frx":0014
      TabIndex        =   4
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Total 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "pocped"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim found As Integer
If Len(vendedor) = 0 Then
   vendedor.SetFocus
   Exit Sub
End If
found = busca_vendedor()
If found = 0 Then
   vendedor = ""
   vendedor.SetFocus
   Exit Sub
End If
found = busca_parame()
If found = 0 Then
   MsgBox "Vendedor no asignado", 48, "Aviso"
   vendedor = ""
   vendedor.SetFocus
   Exit Sub
End If

Frame3.Visible = False
cerrar_dataa
found = crear_temporal_pocket()
If found = 0 Then
   borrar_temporal
End If
found = seleccionar_pocket()
SQL_pedido   'visualizar sus pedidos por vendedor
found = suma_detalle()
found = ir_ultimo()
End Sub
Function crear_temporal_pocket()
On Error GoTo cmd2_err
borrar_archivo globaldir & "\_po" & vendedor & ".dbf"
borrar_archivo globaldir & "\_po" & vendedor & ".cdx"
FileCopy globaldir & "\tdetalle.dbf", globaldir & "\" & "_po" & vendedor & ".dbf"
FileCopy globaldir & "\tdetalle.cdx", globaldir & "\" & "_po" & vendedor & ".cdx"
crear_temporal_pocket = 1
Exit Function
cmd2_err:
Exit Function
End Function
Function seleccionar_pocket()
Dim i As Integer
Dim mytabley As Table
Dim mytablex As Table
Set mytabley = mydbxglo.OpenTable("_po" & vendedor)
mytabley.Index = "tdetalle"
xnueo:
mytabley.Seek "=", "01", "PO", "001", vendedor
If Not mytabley.NoMatch Then
   mytabley.Delete
   GoTo xnueo
End If
Set mytablex = mydbxglo.OpenTable("dproform")
mytablex.Index = "tdetalle"
mytablex.Seek "=", "01", "PO", "001", vendedor
If Not mytablex.NoMatch Then
   Do
     If mytablex.EOF Then Exit Do
     If "" & mytablex.Fields("local") = "01" And "" & mytablex.Fields("tipo") = "PO" And "" & mytablex.Fields("serie") = "001" And "" & mytablex.Fields("numero") = vendedor Then
        '-----------------------------------------
        mytabley.AddNew
        For i = 0 To mytablex.Fields.Count - 1
            mytabley.Fields(i) = mytablex.Fields(i)
        Next i
        mytabley.Fields("local") = "01"
        mytabley.Fields("tipo") = "PO"
        mytabley.Fields("serie") = "001"
        mytabley.Fields("numero") = vendedor
        mytabley.Update
        '-----------------------------------------
        Else: Exit Do
     End If
     mytablex.MoveNext
   Loop
End If
mytablex.Close
mytabley.Close
End Function
Sub SQL_pedido()
Dim buf As String
On Error GoTo cmd1_err
buf = "select * from " & "_po" & vendedor & " where local='01' and tipo='PO' and serie='001' and numero='" & vendedor & "'"
               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldir
               Data2.RecordSource = buf
               Data2.Refresh
               Exit Sub
cmd1_err:
MsgBox "Aviso en slq_pedido " + error$, 48, "Aviso"
Exit Sub
End Sub
Function suma_detalle()
Dim found As Integer
Dim xtotal As Double
found = ir_inicio()
xtotal = 0
Do
If Data2.Recordset.EOF Then Exit Do
xtotal = xtotal + Val("" & Data2.Recordset.Fields("total"))
Data2.Recordset.MoveNext
Loop
Total = Format(xtotal, "0.00")
End Function
Sub borrar_todo()
Dim found As Integer
On Error GoTo cmd90_err
found = ir_inicio()
Do
If Data2.Recordset.EOF Then Exit Do
Data2.Recordset.Delete
Data2.Refresh
Loop
DBGrid2.SetFocus
Exit Sub
cmd90_err:
Exit Sub
End Sub

Function ir_inicio()
On Error GoTo cmd891_err
Data2.Recordset.MoveFirst
ir_inicio = 1
Exit Function
cmd891_err:
Exit Function
End Function

Function ir_ultimo()
On Error GoTo cmd90_err
Data2.Recordset.MoveLast
Exit Function
cmd90_err:
Exit Function
End Function



Function busca_parame()
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("parameca")
mytablex.Index = "caja"
mytablex.Seek "=", vendedor
If Not mytablex.NoMatch Then
   busca_vendedor = 1
End If
mytablex.Close

End Function
Function busca_vendedor()
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("vendedor")
mytablex.Index = "codigo"
mytablex.Seek "=", vendedor
If Not mytablex.NoMatch Then
   busca_vendedor = 1
End If
mytablex.Close
End Function
Private Sub Form_Load()
globaldir = App.Path & "\001d\06"
globaldir = "c:\orion.v5\001D\06"
Set mydbxglo = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
End Sub
Sub borrar_temporal()
On Error GoTo cmd7812_err
mydbxglo.Execute "DELETE FROM " & "_po" & vendedor
Exit Sub
cmd7812_err:
Exit Sub
End Sub
Sub cerrar_dataa()
On Error GoTo cmd43_err
Data2.Recordset.Close
Exit Sub
cmd43_err:
Exit Sub
End Sub

Private Sub Label12_Click()
Dim buf As String
Dim found As Integer
   adiciona_registro
   suma_detalle
   Data2.Refresh
   found = ir_ultimo()
   DBGrid2.SetFocus
   DBGrid2.Col = 1
End Sub
Sub adiciona_registro()
Dim found As Integer
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("producto")
mytablex.Index = "producto"
mytablex.Seek "=", producto
If Not mytablex.NoMatch Then
Data2.Recordset.AddNew
Data2.Recordset.Fields("local") = "01"
Data2.Recordset.Fields("tipo") = "PO"
Data2.Recordset.Fields("SERIE") = "001"
Data2.Recordset.Fields("numero") = vendedor
Data2.Recordset.Fields("linea") = "" & mytablex.Fields("linea")
Data2.Recordset.Fields("producto") = "" & mytablex.Fields("producto")
Data2.Recordset.Fields("descripcio") = "" & mytablex.Fields("descripcio")
Data2.Recordset.Fields("unidad") = "" & mytablex.Fields("unidad")
Data2.Recordset.Fields("factor") = Val("" & mytablex.Fields("facTor"))
Data2.Recordset.Fields("precio") = Val("" & mytablex.Fields("pventa1"))
Data2.Recordset.Fields("igv") = Val("" & mytablex.Fields("igv"))
Data2.Recordset.Fields("total") = Val("" & mytablex.Fields("pventa1"))
Data2.Recordset.Fields("cantidad") = 1
Data2.Recordset.Update
End If
mytablex.Close
End Sub

