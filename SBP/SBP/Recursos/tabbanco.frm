VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form tabbanco 
   BackColor       =   &H00FFFF00&
   Caption         =   "Tabla de Bancos"
   ClientHeight    =   6675
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11580
   StartUpPosition =   1  'CenterOwner
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
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   11415
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "tabbanco.frx":0000
         Height          =   5295
         Left            =   120
         OleObjectBlob   =   "tabbanco.frx":0014
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1080
         Width           =   11055
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5760
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox banco 
      BackColor       =   &H00C0FFFF&
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
      MaxLength       =   6
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox moneda 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox impcheq 
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
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox impdepo 
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
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox saldoini 
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
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox descripcio 
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
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1680
      Width           =   6375
   End
   Begin VB.TextBox codigo 
      BackColor       =   &H00C0FFFF&
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
      MaxLength       =   20
      TabIndex        =   1
      Top             =   960
      Width           =   2895
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
      Picture         =   "tabbanco.frx":09DF
      Style           =   1  'Graphical
      TabIndex        =   13
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
      Picture         =   "tabbanco.frx":1BF1
      Style           =   1  'Graphical
      TabIndex        =   12
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
      Picture         =   "tabbanco.frx":2E03
      Style           =   1  'Graphical
      TabIndex        =   11
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
      Picture         =   "tabbanco.frx":4015
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Imprimir"
      Top             =   0
      Width           =   735
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
      Picture         =   "tabbanco.frx":5227
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
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
      Picture         =   "tabbanco.frx":6439
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Borrar registro"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tabbanco.frx":764B
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Consulta"
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre Banco"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label nbanco 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2280
      TabIndex        =   21
      Top             =   1320
      Width           =   5775
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Banco"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Imp.sobre.Chequ."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Imp.sobre.Depos."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo Inicial"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Denominacion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Num. Cta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   14
      Top             =   960
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
Attribute VB_Name = "tabbanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ajdu1_Click()
If Frame1.Visible = True Then Exit Sub
inicializa
banco = ""
codigo = ""
banco.SetFocus

End Sub

Private Sub banco_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(banco) = 0 Then Exit Sub
found = busca_banco()
If found = 0 Then Exit Sub
codigo.SetFocus
End Sub

Private Sub banco_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
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
banco = ""
inicializa
banco.SetFocus
End Sub

Private Sub buffer_DblClick()
Command1_Click
End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   Frame1.Visible = False
   codigo.SetFocus
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
Frame1.Visible = True
buffer = ""
buffer.SetFocus
Command1_Click
End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   dlo132_Click
   Exit Sub
End If
If Len(banco) = 0 Then
   banco.SetFocus
   Exit Sub
End If
found = busca_banco()
If found = 0 Then
   banco.SetFocus
   Exit Sub
End If
If Len(codigo) = 0 Then
   codigo.SetFocus
   Exit Sub
End If
found = busca_registro()
If found = 0 Then
   inicializa
End If
descripcio.SetFocus
End Sub

Private Sub codigo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   If Len(banco) = 0 Then
      MsgBox "Ingrese una Codigo Valido banco", 48, "Aviso"
      banco.SetFocus
      Exit Sub
   End If
   consulta_datos
End If
If KeyCode = &H26 Then
   banco.SetFocus
   Exit Sub
End If

End Sub

Private Sub Command1_Click()
Dim buf As String
If opcion1 = "1" Then
   If Len(buffer) = 0 Then
      buf = "select Descripcio,Banco from banco "
   Else
   buf = "select Descripcio,Banco from banco where " & Combo1 & " like '" & buffer & "%'"
   End If
End If
If opcion1 = "2" Then
   If Len(buffer) = 0 Then
   buf = "select Descripcio,Cuenta,Banco from tabbanco where banco='" & banco & "'"
   Else
   buf = "select Descripcio,Cuenta,Banco from Tabbanco where banco='" & banco & "' and " & Combo1 & " like '" & buffer & "%'"
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
               dbGrid1.Columns(0).Width = 4000
               dbGrid1.Columns(1).Width = 2000
               dbGrid1.SetFocus

End Sub



Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
   If opcion1 = "1" Then
   banco = dbGrid1.Columns(1)
   codigo = ""
   descripcio = ""
   impdepo = ""
   impcheq = ""
   saldoini = ""
   'banco = ""
   moneda.ListIndex = 0

   Frame1.Visible = False
   banco.SetFocus
   banco_KeyPress 13
   End If
   If opcion1 = "2" Then
   banco = dbGrid1.Columns(2)
   codigo = dbGrid1.Columns(1)
   Frame1.Visible = False
   codigo.SetFocus
   codigo_KeyPress 13
   End If
End If
End Sub


Private Sub descripcio_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
saldoini.SetFocus

End Sub



Private Sub descripcio_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   banco.SetFocus
   Exit Sub
End If

End Sub

Private Sub djuer1_Click()
If Frame1.Visible = True Then Exit Sub
reporgen.NAMETABLA = "tabbanco"
reporgen.Show 1

End Sub

Private Sub dlo132_Click()
If Frame1.Visible = True Then
   If opcion1 = "1" Then
   Frame1.Visible = False
   banco.SetFocus
   Exit Sub
   End If
   If opcion1 = "2" Then
   Frame1.Visible = False
   codigo.SetFocus
   Exit Sub
   End If
End If
tabbanco.Hide
Unload tabbanco
End Sub



Private Sub Form_Load()
moneda.AddItem "S"
moneda.AddItem "D"
moneda.ListIndex = 0
End Sub
Sub inicializa()
nbanco = ""
descripcio = ""
impdepo = ""
impcheq = ""
saldoini = ""
'banco = ""
moneda.ListIndex = 0
End Sub
Function borra_registro()

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("tabbanco")
mytablex.Index = "tabbanco"
mytablex.Seek "=", banco, codigo
If Not mytablex.NoMatch Then
   If MsgBox("Desea Borra el registro", 1, "Aviso") = "1" Then
      mytablex.Delete
      borra_registro = 1
   End If
End If
'------------------------------------- ------------
mytablex.Close
 

End Function
Function busca_registro()

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("tabbanco")
mytablex.Index = "tabbanco"
mytablex.Seek "=", banco, codigo
If Not mytablex.NoMatch Then
   pone_registro mytablex
   busca_registro = 1
End If
'------------------------------------- ------------
mytablex.Close
 
End Function
Sub pone_registro(mytablex As Table)
saldoini = "" & mytablex.Fields("saldoini")
codigo = "" & mytablex.Fields("cuenta")
impdepo = "" & mytablex.Fields("impdepo")
impcheq = "" & mytablex.Fields("impcheq")
descripcio = "" & mytablex.Fields("descripcio")
moneda.ListIndex = 0
If "" & mytablex.Fields("moneda") = "D" Then
   moneda.ListIndex = 1
End If
End Sub
Sub grabando(mytablex As Table)
mytablex.Fields("moneda") = moneda
mytablex.Fields("banco") = banco
mytablex.Fields("cuenta") = codigo
mytablex.Fields("descripcio") = descripcio
mytablex.Fields("impcheq") = Val(impcheq)
mytablex.Fields("impdepo") = Val(impdepo)
mytablex.Fields("saldoini") = Val(saldoini)
End Sub

Private Sub grba1_Click()
Dim found As Integer
If Frame1.Visible = True Then Exit Sub
found = grabar()
If found = 0 Then Exit Sub
codigo.SetFocus

End Sub

Private Sub Image1_Click()
codigo_KeyPress 13
End Sub

Private Sub impcheq_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   impdepo.SetFocus
   Exit Sub
End If

End Sub

Private Sub impdepo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
impcheq.SetFocus

End Sub

Private Sub impdepo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   saldoini.SetFocus
   Exit Sub
End If

End Sub

Private Sub Label1_Click()
cmdSort_Click
End Sub


Function grabar()
Dim found As Integer
Dim mytablex As Table

found = valida()
If found = 0 Then
   MsgBox "Campos invalidos", 48, "Aviso"
   Exit Function
End If

Set mytablex = mydbxglo.OpenTable("tabbanco")
mytablex.Index = "tabbanco"
mytablex.Seek "=", banco, codigo
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
 
End Function

Function valida()
Dim found As Integer
found = busca_banco()
If found = 0 Then
   MsgBox "No existe banco", 48, "Aviso"
   banco.SetFocus
   Exit Function
End If
If Len(codigo) = 0 Then
   codigo.SetFocus
   Exit Function
End If
If Len(descripcio) = 0 Then
   descripcio.SetFocus
   Exit Function
End If
valida = 1
End Function

Private Sub moneda_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   impcheq.SetFocus
   Exit Sub
End If

End Sub

Private Sub saldoini_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
impdepo.SetFocus

End Sub

Private Sub saldoini_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   descripcio.SetFocus
   Exit Sub
End If

End Sub
Function busca_banco()

Dim mytablex As Table
nbanco = ""

Set mytablex = mydbxglo.OpenTable("banco")
mytablex.Index = "banco"
mytablex.Seek "=", banco
If Not mytablex.NoMatch Then
      busca_banco = 1
      nbanco = "" & mytablex.Fields("descripcio")
End If
'------------------------------------- ------------
mytablex.Close
 

End Function
Sub consulta_datos()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "Banco"
Combo1.AddItem "Cuenta"
Combo1.ListIndex = 0

Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "2"
Command1_Click

End Sub
Sub consulta_banco()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "Banco"
Combo1.ListIndex = 0

Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "1"
Command1_Click

End Sub
