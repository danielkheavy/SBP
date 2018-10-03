VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form explorco
   BackColor       =   &H00FFFF00&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Explorador de Pedidos"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   13065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6840
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Consultas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   12735
      Begin VB.TextBox buffer 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "&Ejecutar"
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
         Left            =   8160
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "explorap.frx":0000
         Height          =   5295
         Left            =   120
         OleObjectBlob   =   "explorap.frx":0014
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1080
         Width           =   12495
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Productos"
      Height          =   5295
      Left            =   960
      TabIndex        =   13
      Top             =   960
      Visible         =   0   'False
      Width           =   12015
      Begin VB.CommandButton cmdCancelar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10200
         MaskColor       =   &H00FFFFFF&
         Picture         =   "explorap.frx":09DF
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   4560
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "explorap.frx":118D
         Height          =   4215
         Left            =   120
         OleObjectBlob   =   "explorap.frx":11A1
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   11655
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdGrabar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      Picture         =   "explorap.frx":61BC
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   11640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "explorap.frx":696A
      Height          =   5775
      Left            =   120
      OleObjectBlob   =   "explorap.frx":697E
      TabIndex        =   8
      Top             =   960
      Width           =   12735
   End
   Begin VB.ComboBox estado 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox codigo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      MaxLength       =   11
      TabIndex        =   5
      Text            =   "*"
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox fechaf 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox fechai 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label totald 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   11
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label totals 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   10
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   9
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Menu ldo33 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "explorco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub buffer_DblClick()
Command1_Click
End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   ldo33_Click
   Exit Sub
End If
Command1_Click

End Sub

Private Sub cmdCancelar_Click()
ldo33_Click
End Sub

Private Sub cmdGrabar_Click()
sql_cabeza
End Sub

Private Sub codigo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
consulta_codigo
End If

End Sub

Private Sub Command1_Click()
Dim buf As String
If opcion1 = "1" Then
If Len(buffer) = 0 Then
buf = "select Nombre,Codigo from clientes "
Else
buf = "select Nombre,Codigo from clientes where " & Combo1 & " like '" & buffer & "*'"
End If
End If


               Data1.Connect = "foxpro 2.5;"
               Data1.DatabaseName = globaldir
               Data1.RecordSource = buf
               Data1.Refresh
               If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
                  Data1.Recordset.Close
                  buffer.SetFocus
                  Exit Sub
               End If
               DBGrid1.Columns(0).Width = 4000
               DBGrid1.Columns(1).Width = 2000
               DBGrid1.SetFocus

End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
  
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
   If opcion1 = "1" Then
   codigo = DBGrid1.Columns(1)
   Frame1.Visible = False
   codigo.SetFocus
   
   End If
End If

End Sub

Private Sub DBGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H71 Then  'f2
cambia_estado
End If
If KeyCode = &H70 Then  'f1
consulta_productos
End If

End Sub

Private Sub Form_Load()
fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
fechaf = Format(Now, "dd/mm/yyyy")
estado.Clear
estado.AddItem "*"
estado.AddItem "Pendientes"
estado.AddItem "Despachados"
estado.ListIndex = 0
End Sub

Private Sub ldo33_Click()
If Frame2.Visible = True Then
   Frame2.Visible = False
   DBGrid2.SetFocus
   Exit Sub
End If
If opcion1 = "1" Then
If Frame1.Visible = True Then
   Frame1.Visible = False
   codigo.SetFocus
   Exit Sub
End If
End If
explorap.Hide
Unload explorap
End Sub
Sub sql_cabeza()
On Error GoTo cmd37_err
Dim buf As String
If Not IsDate(fechai) Then Exit Sub
If Not IsDate(fechaf) Then Exit Sub
buf = "select * from deliveri where "
buf = buf & "  fecha>=" & "DateValue('" & fechai & "'" & ")"
buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
If codigo <> "*" Then
   buf = buf & " and codigo like '" & codigo & "'"
End If
If estado = "Pendiente" Then
buf = buf & " and (yausado=null or yausado<>'1')"
End If
If estado = "Despachados" Then
buf = buf & " and yausado='1'"
End If

               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldir
               Data2.RecordSource = buf
               Data2.Refresh
               
               SUMAR_CABEZA
               DBGrid2.SetFocus
               
Exit Sub
cmd37_err:
MsgBox "Error en select " & Error$, 48, "Aviso"
Exit Sub
End Sub
Sub SUMAR_CABEZA()
Dim sdx As Double
Dim sdx1 As Double
sdx = 0
sdx1 = 0
ir_inicio
Do
If Data2.Recordset.EOF Then Exit Do
If "" & Data2.Recordset.Fields("moneda") = "S" Then
   sdx = sdx + Val("" & Data2.Recordset.Fields("total"))
End If
If "" & Data2.Recordset.Fields("moneda") = "D" Then
   sdx1 = sdx1 + Val("" & Data2.Recordset.Fields("total"))
End If
Data2.Recordset.MoveNext
Loop
totals = Format(sdx, "0.00")
totald = Format(sdx, "0.00")
End Sub
Sub ir_inicio()
On Error GoTo cmd4_err
Data2.Recordset.movefist
Exit Sub
cmd4_err:
Exit Sub
End Sub
Sub consulta_codigo()
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Telefono"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "1"
Command1_Click

End Sub
Sub cambia_estado()
If "" & Data2.Recordset.Fields("yausado") = "1" Then
   Data2.Recordset.Edit
   Data2.Recordset.Fields("yausado") = "0"
   Data2.Recordset.Update
   Exit Sub
End If
If "" & Data2.Recordset.Fields("yausado") = "0" Or "" & Data2.Recordset.Fields("yausado") = "" Then
   Data2.Recordset.Edit
   Data2.Recordset.Fields("yausado") = "1"
   Data2.Recordset.Update
   Exit Sub
End If

End Sub
Sub consulta_productos()
On Error GoTo cmd39_err
Dim buf As String
buf = "select * from ddeliver where "
buf = buf & " tipo='" & "" & Data2.Recordset.Fields("tipo") & "'"
buf = buf & " and serie='" & "" & Data2.Recordset.Fields("serie") & "'"
buf = buf & " and numero='" & "" & Data2.Recordset.Fields("numero") & "'"
               Data3.Connect = "foxpro 2.5;"
               Data3.DatabaseName = globaldir
               Data3.RecordSource = buf
               Frame2.Visible = True
               Data3.Refresh
               DBGrid3.SetFocus
Exit Sub
cmd39_err:
Exit Sub
End Sub

