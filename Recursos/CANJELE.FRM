VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form canjele 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Canje Facturas x Letras"
   ClientHeight    =   7530
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Height          =   2775
      Left            =   120
      OleObjectBlob   =   "canjele.frx":0000
      TabIndex        =   27
      Top             =   2520
      Width           =   9735
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      MaskColor       =   &H00E0E0E0&
      Picture         =   "canjele.frx":09D3
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox moneda 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   25
      Top             =   2040
      Width           =   1575
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
      Left            =   10080
      TabIndex        =   19
      Top             =   1440
      Visible         =   0   'False
      Width           =   11415
      Begin VB.ComboBox Combo2 
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
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5160
         Visible         =   0   'False
         Width           =   1140
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
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
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
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
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
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "canjele.frx":1BE5
         Height          =   5295
         Left            =   120
         OleObjectBlob   =   "canjele.frx":1BF9
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1080
         Width           =   11055
      End
   End
   Begin VB.TextBox girador 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      MaxLength       =   11
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox aceptante 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      MaxLength       =   11
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox importe 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      MaxLength       =   10
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox fechaf 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox fechai 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox codigo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "canjele.frx":25C4
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Grabar registro"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdAddEntry 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      Picture         =   "canjele.frx":37D6
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Nuevo registro"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      MaskColor       =   &H00E0E0E0&
      Picture         =   "canjele.frx":49E8
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Ayuda"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      MaskColor       =   &H00E0E0E0&
      Picture         =   "canjele.frx":5BFA
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Imprimir"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      MaskColor       =   &H00E0E0E0&
      Picture         =   "canjele.frx":6E0C
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Borrar registro"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "canjele.frx":801E
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Consulta"
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Girador"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   18
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aceptante"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   17
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Importe"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   15
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Vencimiento"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Emision"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nro.Letra"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   2175
   End
   Begin VB.Menu ajdu1 
      Caption         =   "&Add"
   End
   Begin VB.Menu bo712 
      Caption         =   "&Borrar"
   End
   Begin VB.Menu grba1 
      Caption         =   "&Grabar"
   End
   Begin VB.Menu djuer1 
      Caption         =   "&Reporte"
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "canjele"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub aceptante_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(aceptante) = 0 Then Exit Sub
found = busca_codigo("" & aceptante)
If found = 0 Then
   MsgBox "No existe aceptante", 48, "Aviso"
   Exit Sub
End If
girador.SetFocus

End Sub

Private Sub aceptante_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   paridad.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   consulta_aceptante
End If

End Sub

Private Sub agencia_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
refactura.SetFocus

End Sub

Private Sub agencia_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   seccion.SetFocus
   Exit Sub
End If

End Sub

Private Sub ajdu1_Click()
If Frame1.Visible = True Then Exit Sub
inicializa
codigo = ""
codigo.SetFocus

End Sub

Private Sub banco_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(banco) > 0 Then
   found = busca_banco()
   If found = 0 Then
      MsgBox "No existe Banco", 48, "Aviso"
      Exit Sub
   End If
End If
seccion.SetFocus

End Sub

Private Sub banco_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   girador.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   consulta_banco
End If


End Sub

Private Sub bo712_Click()
Dim found As Integer
If Frame1.Visible = True Then Exit Sub
found = busca_registro()
If found = 0 Then
   MsgBox "No existe registro", 48, "Aviso"
   Exit Sub
End If
found = borra_registro()
If found = 0 Then Exit Sub
MsgBox "Ok,Registro Borrado", 48, "Aviso"
codigo = ""
inicializa
codigo.SetFocus
End Sub

Private Sub buffer_DblClick()
Command1_Click
End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   dlo132_Click
   Exit Sub
End If
Command1_Click

End Sub

Private Sub cmdAddEntry_Click()
ajdu1_Click
End Sub

Private Sub cmdDelete_Click()
bo712_Click
End Sub

Private Sub cmdExit_Click()
dlo132_Click

End Sub

Private Sub cmdPrint_Click()
djuer1_Click
End Sub

Private Sub cmdSave_Click()
grba1_Click
End Sub

Private Sub cmdSort_Click()
cerrar_data1
Combo1.Clear
Combo1.AddItem "Aceptante"
Combo1.AddItem "Letra"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "1"
End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   dlo132_Click
   Exit Sub
End If
If Len(codigo) = 0 Then Exit Sub
found = busca_registro()
If found = 0 Then
   inicializa
End If
fechai.SetFocus
End Sub

Private Sub codigo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
cmdSort_Click
End If
End Sub

Private Sub Command1_Click()
Dim buf As String
   If opcion1 = "2" Or opcion1 = "3" Then
      If Len(buffer) = 0 Then
      buf = "select Nombre,Codigo from clientes "
      Else
      buf = "select Nombre,Codigo from clientes where " & Combo1 & " like '" & buffer & "*'"
      End If
      If Combo2.ListIndex = 1 Then
         buf = buf & " order by " & Combo1
      End If
   End If
   If opcion1 = "1" Then
      If Len(buffer) = 0 Then
      buf = "select Aceptante,letra from letrac "
      Else
      buf = "select Aceptante,letra from letrac where " & Combo1 & " like '" & buffer & "*'"
      End If
      If Combo2.ListIndex = 1 Then
         buf = buf & " order by " & Combo1
      End If
   End If
   If opcion1 = "4" Then
      If Len(buffer) = 0 Then
      buf = "select Descripcio,Banco from Banco "
      Else
      buf = "select Descripcio,Banco from Banco where " & Combo1 & " like '" & buffer & "*'"
      End If
      If Combo2.ListIndex = 1 Then
         buf = buf & " order by " & Combo1
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
               If opcion1 = "1" Or opcion2 = "2" Or opcion1 = "3" Or opcion1 = "4" Then
                  DBGrid1.Columns(0).Width = 4000
                  DBGrid1.Columns(1).Width = 2000
               End If
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
   codigo_KeyPress 13
End If
If opcion1 = "2" Then
   aceptante = DBGrid1.Columns(1)
   Frame1.Visible = False
   aceptante.SetFocus
   aceptante_KeyPress 13
End If
If opcion1 = "3" Then
   girador = DBGrid1.Columns(1)
   Frame1.Visible = False
   girador.SetFocus
   girador_KeyPress 13
End If
If opcion1 = "4" Then
   banco = DBGrid1.Columns(1)
   Frame1.Visible = False
   banco.SetFocus
   banco_KeyPress 13
End If


End If
End Sub


Private Sub descripcio_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub


End Sub



Private Sub descripcio_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   codigo.SetFocus
   Exit Sub
End If

End Sub

Private Sub djuer1_Click()
If Frame1.Visible = True Then Exit Sub
reporgen.NAMETABLA = "letrac"
reporgen.Show 1

End Sub

Private Sub dlo132_Click()
If Frame1.Visible = True Then
   If opcion1 = "1" Then
      Frame1.Visible = False
      codigo.SetFocus
      Exit Sub
   End If
   If opcion1 = "2" Then
      Frame1.Visible = False
      aceptante.SetFocus
      Exit Sub
   End If
   If opcion1 = "3" Then
      Frame1.Visible = False
      girador.SetFocus
      Exit Sub
   End If
   If opcion1 = "4" Then
      Frame1.Visible = False
      banco.SetFocus
      Exit Sub
   End If
   Exit Sub
End If
canjele.Hide
Unload canjele
End Sub



Private Sub estado_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   observa.SetFocus
   Exit Sub
End If

End Sub

Private Sub fechaf_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Len(fechaf) = 0 Then
   fechaf = Format(Now, "dd/mm/yyyy")
End If
If Not IsDate(fechaf) Then Exit Sub
If Len(fechaf) <> 10 Then Exit Sub
importe.SetFocus

End Sub

Private Sub fechaf_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   fechai.SetFocus
   Exit Sub
End If

End Sub

Private Sub fechai_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Len(fechai) = 0 Then
   fechai = Format(Now, "dd/mm/yyyy")
End If
If Not IsDate(fechai) Then Exit Sub
If Len(fechai) <> 10 Then Exit Sub
fechaf.SetFocus
End Sub

Private Sub fechai_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   codigo.SetFocus
   Exit Sub
End If

End Sub

Private Sub Form_Load()

Dim mydbx As Database
Dim mytablex As Table
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("carsec")
Do
If mytablex.EOF Then Exit Do
seccion.AddItem "" & mytablex.Fields("carsec")
mytablex.MoveNext
Loop
mytablex.Close
mydbx.Close
seccion.ListIndex = 0

moneda.Clear
moneda.AddItem "S"
moneda.AddItem "D"
moneda.ListIndex = 0

Combo1.Clear
Combo1.AddItem "Aceptante"
Combo1.AddItem "letra"
Combo1.ListIndex = 0
End Sub
Sub inicializa()
fechai = ""
fechaf = ""
importe = ""
moneda.ListIndex = 0
paridad = ""
aceptante = ""
girador = ""
banco = ""
seccion.ListIndex = 0
agencia = ""
refactura = ""
observa = ""
estado = ""

End Sub
Function borra_registro()
Dim mydbx As Database
Dim mytablex As Table
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("letrac")
mytablex.Index = "letra"
mytablex.Seek "=", codigo
If Not mytablex.NoMatch Then
   If MsgBox("Desea Borra el registro", 1, "Aviso") = "1" Then
      mytablex.Delete
      borra_registro = 1
   End If
End If
'------------------------------------- ------------
mytablex.Close
mydbx.Close

End Function
Function busca_registro()
Dim mydbx As Database
Dim mytablex As Table
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("letrac")
mytablex.Index = "letra"
mytablex.Seek "=", codigo
If Not mytablex.NoMatch Then
   pone_registro mytablex
   busca_registro = 1
End If
'------------------------------------- ------------
mytablex.Close
mydbx.Close
End Function
Sub pone_registro(mytablex As Table)
codigo = "" & mytablex.Fields("letra")
fechai = "" & mytablex.Fields("fechai")
fechaf = "" & mytablex.Fields("fechaf")
importe = "" & mytablex.Fields("importe")
moneda = "" & mytablex.Fields("moneda")
paridad = "" & mytablex.Fields("paridad")
aceptante = "" & mytablex.Fields("aceptante")
girador = "" & mytablex.Fields("girador")
banco = "" & mytablex.Fields("banco")
seccion = "" & mytablex.Fields("seccion")
agencia = "" & mytablex.Fields("agencia")
refactura = "" & mytablex.Fields("refactura")
observa = "" & mytablex.Fields("observa")
estado = "" & mytablex.Fields("estado")
End Sub
Sub grabando(mytablex As Table)

mytablex.Fields("letra") = codigo
mytablex.Fields("fechai") = Format(fechai, "dd/mm/yyyy")
mytablex.Fields("fechaf") = Format(fechai, "dd/mm/yyyy")
mytablex.Fields("importe") = Val(importe)
mytablex.Fields("moneda") = moneda
mytablex.Fields("paridad") = Val(paridad)
mytablex.Fields("aceptante") = aceptante
mytablex.Fields("girador") = girador
mytablex.Fields("banco") = banco
mytablex.Fields("seccion") = seccion
mytablex.Fields("agencia") = agencia
mytablex.Fields("refactura") = refactura
mytablex.Fields("observa") = observa
mytablex.Fields("estado") = estado
End Sub

Private Sub girador_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(girador) = 0 Then Exit Sub
found = busca_codigo("" & girador)
If found = 0 Then
   MsgBox "No existe Girador", 48, "Aviso"
   Exit Sub
End If
banco.SetFocus

End Sub

Private Sub girador_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   aceptante.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   consulta_girador
End If


End Sub

Private Sub grba1_Click()
Dim found As Integer
If Frame1.Visible = True Then Exit Sub
found = grabar()
If found = 0 Then Exit Sub
codigo.SetFocus

End Sub

Private Sub importe_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
moneda.SetFocus
End Sub

Private Sub importe_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   fechaf.SetFocus
   Exit Sub
End If

End Sub

Private Sub Label1_Click()
cmdSort_Click
End Sub


Function grabar()
Dim found As Integer
Dim mytablex As Table
Dim mydbx As Database
found = valida()
If found = 0 Then
   MsgBox "Campos invalidos", 48, "Aviso"
   Exit Function
End If
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("letrac")
mytablex.Index = "letra"
mytablex.Seek "=", codigo
If mytablex.NoMatch Then
   mytablex.AddNew
   grabando mytablex
   mytablex.Update
   grabar = 1
End If
If Not mytablex.NoMatch Then
   If MsgBox("Desea Reescribir?", 1, "Aviso") = 1 Then
   mytablex.Edit
   grabando mytablex
   mytablex.Update
   grabar = 1
   End If
End If
'------------------------------------- ------------
mytablex.Close
mydbx.Close
End Function

Function valida()
Dim found As Integer
If Len(codigo) = 0 Then
   codigo.SetFocus
   Exit Function
End If
If Not IsDate(fechai) Or Len(fechai) <> 10 Then
   fechai = ""
   fechai.SetFocus
   Exit Function
End If
If Not IsDate(fechaf) Or Len(fechaf) <> 10 Then
   fechaf = ""
   fechaf.SetFocus
   Exit Function
End If
If Len(aceptante) = 0 Then
   aceptante.SetFocus
   Exit Function
End If
found = busca_codigo("" & aceptante)
If found = 0 Then
   MsgBox "No existe aceptante", 48, "Aviso"
   aceptante.SetFocus
   Exit Function
End If
If Len(girador) = 0 Then
   girador.SetFocus
   Exit Function
End If
found = busca_codigo("" & girador)
If found = 0 Then
   MsgBox "No existe girador", 48, "Aviso"
   girador.SetFocus
   Exit Function
End If
If Len(banco) > 0 Then
   found = busca_banco()
   If found = 0 Then
      MsgBox "No existe Banco", 48, "Aviso"
      banco = ""
      banco.SetFocus
      Exit Function
   End If
End If

valida = 1
End Function

Private Sub moneda_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
paridad.SetFocus

End Sub

Private Sub moneda_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   importe.SetFocus
   Exit Sub
End If

End Sub

Private Sub observa_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
estado.SetFocus

End Sub

Private Sub observa_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   refactura.SetFocus
   Exit Sub
End If

End Sub

Private Sub paridad_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
aceptante.SetFocus

End Sub

Private Sub paridad_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   moneda.SetFocus
   Exit Sub
End If

End Sub

Private Sub refactura_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
observa.SetFocus

End Sub

Private Sub refactura_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   agencia.SetFocus
   Exit Sub
End If

End Sub

Private Sub seccion_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
agencia.SetFocus
End Sub

Private Sub seccion_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   banco.SetFocus
   Exit Sub
End If

End Sub
Function busca_codigo(buf As String)
Dim mydbx As Database
Dim mytablex As Table
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("clientes")
mytablex.Index = "codigo"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   busca_codigo = 1
End If
'------------------------------------- ------------
mytablex.Close
mydbx.Close

End Function
Function busca_banco()
Dim mydbx As Database
Dim mytablex As Table
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("banco")
mytablex.Index = "banco"
mytablex.Seek "=", banco
If Not mytablex.NoMatch Then
   busca_banco = 1
End If
'------------------------------------- ------------
mytablex.Close
mydbx.Close

End Function
Sub consulta_aceptante()
cerrar_data1
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Codigo"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "2"

End Sub
Sub consulta_girador()
cerrar_data1
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Codigo"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "3"

End Sub
Sub consulta_banco()
cerrar_data1
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "Banco"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "4"

End Sub
Sub cerrar_data1()
On Error GoTo cmd17_err
Data1.Recordset.Close
Exit Sub
cmd17_err:
Exit Sub
End Sub

