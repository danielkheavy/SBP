VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form tcuagreg 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuadres generales Oficina"
   ClientHeight    =   6360
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
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
      Height          =   6255
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   10695
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         Left            =   8040
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "Tcuadreg.frx":0000
         Height          =   5055
         Left            =   120
         OleObjectBlob   =   "Tcuadreg.frx":0014
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1080
         Width           =   10455
      End
   End
   Begin VB.Data Data3 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4680
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid3 
      Bindings        =   "Tcuadreg.frx":09DF
      Height          =   1815
      Left            =   120
      OleObjectBlob   =   "Tcuadreg.frx":09F3
      TabIndex        =   19
      Top             =   4440
      Width           =   10575
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5400
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "Tcuadreg.frx":1F66
      Height          =   2295
      Left            =   120
      OleObjectBlob   =   "Tcuadreg.frx":1F7A
      TabIndex        =   13
      Top             =   1680
      Width           =   10575
   End
   Begin VB.TextBox paridad 
      Height          =   285
      Left            =   6720
      MaxLength       =   10
      TabIndex        =   11
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox descripcio 
      Height          =   285
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   9
      Top             =   1080
      Width           =   4215
   End
   Begin VB.TextBox codigo 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   1140
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
      Picture         =   "Tcuadreg.frx":3805
      Style           =   1  'Graphical
      TabIndex        =   7
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
      Picture         =   "Tcuadreg.frx":4A17
      Style           =   1  'Graphical
      TabIndex        =   6
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
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "Tcuadreg.frx":5C29
      Style           =   1  'Graphical
      TabIndex        =   5
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
      Picture         =   "Tcuadreg.frx":6E3B
      Style           =   1  'Graphical
      TabIndex        =   4
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
      Picture         =   "Tcuadreg.frx":804D
      Style           =   1  'Graphical
      TabIndex        =   3
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
      Picture         =   "Tcuadreg.frx":925F
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Borrar registro"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   3600
      MaskColor       =   &H00E0E0E0&
      Picture         =   "Tcuadreg.frx":A471
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Consulta"
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Entrega"
      Height          =   255
      Left            =   8640
      TabIndex        =   25
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Neto"
      Height          =   255
      Left            =   6360
      TabIndex        =   24
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Documentos"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Efectivo"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Egresos"
      Height          =   255
      Left            =   4440
      TabIndex        =   21
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ingresos"
      Height          =   255
      Left            =   2400
      TabIndex        =   20
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      Height          =   255
      Left            =   2400
      TabIndex        =   18
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label label53 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "L4"
      Height          =   255
      Left            =   9000
      TabIndex        =   17
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label label52 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "L3"
      Height          =   255
      Left            =   7440
      TabIndex        =   16
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label label51 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "L2"
      Height          =   255
      Left            =   5880
      TabIndex        =   15
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label label50 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "L1"
      Height          =   255
      Left            =   4200
      TabIndex        =   14
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "T/Cambio"
      Height          =   255
      Left            =   5520
      TabIndex        =   12
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Observacion"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1215
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
Attribute VB_Name = "tcuagreg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub ajdu1_Click()
If Frame1.Visible = True Then Exit Sub
inicializa
codigo = ""
codigo.SetFocus

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
If Len(codigo) = 0 Then
   codigo = Format(Now, "dd/mm/yyyy")
   Exit Sub
End If
If Len(codigo) <> 10 Then Exit Sub
If Not IsDate(codigo) Then Exit Sub
found = busca_registro()
If found = 0 Then
   inicializa
End If
sql_cuadre01
sql_cuadre02
descripcio.SetFocus
End Sub
Sub sql_cuadre01()
Dim buf As String
buf = "select * from cuadre01 where "
buf = buf & " fecha=" & "DateValue('" & codigo & "'" & ")"
               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldir
               Data2.RecordSource = buf
               Data2.Refresh
End Sub
Sub sql_cuadre02()
Dim buf As String
buf = "select * from cuadre02 where "
buf = buf & " fecha=" & "DateValue('" & codigo & "'" & ")"
               Data3.Connect = "foxpro 2.5;"
               Data3.DatabaseName = globaldir
               Data3.RecordSource = buf
               Data3.Refresh

End Sub

Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
cmdSort_Click
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
codigo.SetFocus
End Sub

Private Sub Command1_Click()
Dim buf As String
If Len(buffer) = 0 Then
buf = "select Descripcio,Fecha from cuadrege "
Else
buf = "select Descripcio,Fecha from cuadrege where " & Combo1 & " like '" & buffer & "%'"
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
   codigo = dbGrid1.Columns(1)
   Frame1.Visible = False
   codigo.SetFocus
   codigo_KeyPress 13
End If
End Sub



Private Sub DBGrid3_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If ColIndex < 6 And ColIndex > 7 Then
       Cancel = True
       Exit Sub
End If
Select Case ColIndex
       
       Case 7, 8
            If Len("" & dbGrid3.Columns(0)) = 0 Then
               Cancel = True
               Exit Sub
            End If
       
End Select

End Sub

Private Sub DBGrid3_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Select Case ColIndex
       Case 6
            If Not IsNumeric(dbGrid3.Columns(6)) Then
               Cancel = True
               Exit Sub
            End If
       Case 7
            If Not IsNumeric(dbGrid3.Columns(7)) Then
               Cancel = True
               Exit Sub
            End If
            
End Select


End Sub

Private Sub descripcio_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
paridad.SetFocus


End Sub



Private Sub descripcio_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   codigo.SetFocus
   Exit Sub
End If

End Sub


Private Sub djuer1_Click()
If Frame1.Visible = True Then Exit Sub
reporgen.NAMETABLA = "cuadrege"
reporgen.Show 1

End Sub

Private Sub dlo132_Click()
If Frame1.Visible = True Then
   Frame1.Visible = False
   Exit Sub
End If
tcuagreg.Hide
Unload tcuagreg
End Sub




Private Sub Form_Activate()

Dim mytablex As Table


Set mytablex = mydbxglo.OpenTable("grupos")
mytablex.Index = "grupos"
mytablex.Seek "=", "01"
If Not mytablex.NoMatch Then
   Label50 = "" & mytablex.Fields("l1")
   Label51 = "" & mytablex.Fields("l2")
   Label52 = "" & mytablex.Fields("l3")
   Label53 = "" & mytablex.Fields("l4")
   
End If
'------------------------------------- ------------
mytablex.Close
 

End Sub

Private Sub Form_Load()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "Grupos"
Combo1.ListIndex = 0
End Sub
Sub inicializa()
descripcio = ""
paridad = ""
End Sub
Function borra_registro()

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("cuadrege")
mytablex.Index = "cuadrege"
mytablex.Seek "=", codigo
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

Set mytablex = mydbxglo.OpenTable("cuadrege")
mytablex.Index = "cuadrege"
mytablex.Seek "=", codigo
If Not mytablex.NoMatch Then
   pone_registro mytablex
   busca_registro = 1
End If
'------------------------------------- ------------
mytablex.Close
 
End Function
Sub pone_registro(mytablex As Table)
codigo = "" & mytablex.Fields("fecha")
descripcio = "" & mytablex.Fields("descripcio")
paridad = "" & mytablex.Fields("paridad")
End Sub
Sub grabando(mytablex As Table)
mytablex.Fields("fecha") = codigo
mytablex.Fields("descripcio") = descripcio
mytablex.Fields("paridad") = Val(paridad)
End Sub

Private Sub grba1_Click()
Dim found As Integer
If Frame1.Visible = True Then Exit Sub
found = grabar()
If found = 0 Then Exit Sub
codigo.SetFocus

End Sub

Private Sub grupos_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
codigo.SetFocus
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

Set mytablex = mydbxglo.OpenTable("cuadrege")
mytablex.Index = "cuadrege"
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
 
End Function

Function valida()
If Len(codigo) = 0 Then
   codigo.SetFocus
   Exit Function
End If
If Len(codigo) <> 10 Then
   codigo.SetFocus
   Exit Function
End If
If Not IsDate(codigo) Then
   codigo.SetFocus
   Exit Function
End If
If Len(descripcio) = 0 Then
   descripcio.SetFocus
   Exit Function
End If
valida = 1
End Function

Private Sub Label18_Click()
Generar_sumas
End Sub

Private Sub label51_Click()
If MsgBox("Se borrar acumulados,Continua...", 1, "Aviso") <> 1 Then Exit Sub
borrar_acumulados
Generar_sumas
sql_cuadre01
sql_cuadre02

End Sub

Private Sub paridad_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub


End Sub

Sub Generar_sumas()
Dim mytableb As Snapshot
Dim mytablez As Table
Dim mytablex As Table

Dim buf As String


Set mytablex = mydbxglo.OpenTable("cuadre01")
mytablex.Index = "cuadre01"
Set mytablez = mydbxglo.OpenTable("cuadre02")
mytablez.Index = "cuadre02"
buf = "select * from factura where "
buf = buf & " fecha=" & "DateValue('" & codigo & "'" & ")"
buf = buf & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or "
buf = buf & " acu='A' or acu='B' or acu='C' or acu='D' or acu='G'"
buf = buf & " or acu='E' or acu='F' or acu='N' or acu='O') "
buf = buf & " and estado='2'"
Set mytableb = mydbxglo.CreateSnapshot(buf)
Do
If mytableb.EOF Then Exit Do
          mytablex.Seek "=", codigo, "" & mytableb.Fields("acu")
          If mytablex.NoMatch Then
             mytablex.AddNew
             mytablex.Fields("tipo") = "" & mytableb.Fields("acu")
             mytablex.Fields("descripcio") = pone_descripcio("" & mytableb.Fields("acu"))
             mytablex.Fields("fecha") = Format(codigo, "dd/mm/yyyy")
             graba_campos mytablex, mytableb
             mytablex.Update
          End If
          If Not mytablex.NoMatch Then
             mytablex.Edit
             graba_campos mytablex, mytableb
             mytablex.Update
          End If

mytableb.MoveNext
Loop
mytableb.Close
'ahora los recibos
buf = "select * from recibo where "
buf = buf & " fecha=" & "DateValue('" & codigo & "'" & ")"
buf = buf & " and estado='2'"
Set mytableb = mydbxglo.CreateSnapshot(buf)
Do
If mytableb.EOF Then Exit Do
          mytablex.Seek "=", codigo, "" & mytableb.Fields("acu")
          If mytablex.NoMatch Then
             mytablex.AddNew
             mytablex.Fields("tipo") = "" & mytableb.Fields("acu")
             mytablex.Fields("descripcio") = pone_descripcio("" & mytableb.Fields("acu"))
             mytablex.Fields("fecha") = Format(codigo, "dd/mm/yyyy")
             graba_campos mytablex, mytableb
             mytablex.Update
          End If
          If Not mytablex.NoMatch Then
             mytablex.Edit
             graba_campos mytablex, mytableb
             mytablex.Update
          End If
mytableb.MoveNext
Loop
mytableb.Close
'ahora el efectivo
buf = "select * from fpagov where "
buf = buf & " fecha=" & "DateValue('" & codigo & "'" & ")"
buf = buf & " and estado='2'"
Set mytableb = mydbxglo.CreateSnapshot(buf)
Do
If mytableb.EOF Then Exit Do
              mytablez.Seek "=", codigo, "" & mytableb.Fields("acufp")
              If mytablez.NoMatch Then
                 mytablez.AddNew
                 mytablez.Fields("tipo") = "" & mytableb.Fields("acufp")
                 mytablez.Fields("descripcio") = pone_fpago("" & mytableb.Fields("acufp"))
                 mytablez.Fields("fecha") = Format(codigo, "dd/mm/yyyy")
                 graba_campos1 mytablez, mytableb
                 mytablez.Update
              End If
              If Not mytablez.NoMatch Then
                 mytablez.Edit
                 graba_campos1 mytablez, mytableb
                 mytablez.Update
              End If
           
mytableb.MoveNext
Loop
mytableb.Close
mytablex.Close
mytablez.Close
 
End Sub
Sub graba_campos(mytablex As Table, mytableb As Snapshot)
Dim ssdx As Double
If "" & mytableb.Fields("moneda") = "S" Then
ssdx = Val("" & mytablex.Fields("tts")) + Val("" & mytableb.Fields("Total"))
mytablex.Fields("tts") = ssdx
End If
If "" & mytableb.Fields("moneda") = "D" Then
ssdx = Val("" & mytablex.Fields("ttd")) + Val("" & mytableb.Fields("Total"))
mytablex.Fields("ttd") = ssdx
End If


             If Val("" & mytableb.Fields("c1")) > 0 Then
                If "" & mytableb.Fields("moneda") = "S" Then
                   ssdx = Val("" & mytablex.Fields("t1s")) + Val("" & mytableb.Fields("c1"))
                   mytablex.Fields("t1s") = ssdx
                End If
                If "" & mytableb.Fields("moneda") = "D" Then
                   ssdx = Val("" & mytablex.Fields("t1d")) + Val("" & mytableb.Fields("c1"))
                   mytablex.Fields("t1d") = ssdx
                End If
             End If
             If Val("" & mytableb.Fields("c2")) > 0 Then
                If "" & mytableb.Fields("moneda") = "S" Then
                   ssdx = Val("" & mytablex.Fields("t2s")) + Val("" & mytableb.Fields("c2"))
                   mytablex.Fields("t2s") = ssdx
                End If
                If "" & mytableb.Fields("moneda") = "D" Then
                   ssdx = Val("" & mytablex.Fields("t2d")) + Val("" & mytableb.Fields("c2"))
                   mytablex.Fields("t2d") = ssdx
                End If
             End If
             If Val("" & mytableb.Fields("c3")) > 0 Then
                If "" & mytableb.Fields("moneda") = "S" Then
                   ssdx = Val("" & mytablex.Fields("t3s")) + Val("" & mytableb.Fields("c3"))
                   mytablex.Fields("t3s") = ssdx
                End If
                If "" & mytableb.Fields("moneda") = "D" Then
                   ssdx = Val("" & mytablex.Fields("t3d")) + Val("" & mytableb.Fields("c3"))
                   mytablex.Fields("t3d") = ssdx
                End If
             End If
             If Val("" & mytableb.Fields("c4")) > 0 Then
                If "" & mytableb.Fields("moneda") = "S" Then
                   ssdx = Val("" & mytablex.Fields("t4s")) + Val("" & mytableb.Fields("c4"))
                   mytablex.Fields("t4s") = ssdx
                End If
                If "" & mytableb.Fields("moneda") = "D" Then
                   ssdx = Val("" & mytablex.Fields("t4d")) + Val("" & mytableb.Fields("c4"))
                   mytablex.Fields("t4d") = ssdx
                End If
             End If
End Sub
Sub graba_campos1(mytablex As Table, mytableb As Snapshot)
Dim ssdx As Double
  Select Case "" & mytableb.Fields("acu")
         Case "A", "B", "C", "D", "G", "E", "F", "W"
                If "" & mytableb.Fields("moneda") = "S" Then
                   ssdx = Val("" & mytablex.Fields("isoles")) + Val("" & mytableb.Fields("total"))
                   mytablex.Fields("isoles") = ssdx
                End If
                If "" & mytableb.Fields("moneda") = "D" Then
                   ssdx = Val("" & mytablex.Fields("idolares")) + Val("" & mytableb.Fields("total"))
                   mytablex.Fields("idolares") = ssdx
                End If
         Case "V", "K", "L", "M", "P", "N", "O", "X"
                If "" & mytableb.Fields("moneda") = "S" Then
                   ssdx = Val("" & mytablex.Fields("esoles")) + Val("" & mytableb.Fields("total"))
                   mytablex.Fields("esoles") = ssdx
                End If
                If "" & mytableb.Fields("moneda") = "D" Then
                   ssdx = Val("" & mytablex.Fields("edolares")) + Val("" & mytableb.Fields("total"))
                   mytablex.Fields("edolares") = ssdx
                End If
 End Select

End Sub
Sub borrar_acumulados()


mydbxglo.Execute "DELETE FROM cuadre01 where  fecha=" & "DateValue('" & codigo & "'" & ")"
mydbxglo.Execute "DELETE FROM cuadre02 where  fecha=" & "DateValue('" & codigo & "'" & ")"
 
End Sub
Function pone_descripcio(buf As String) As String
Select Case buf
       Case "A", "B", "C", "D"
            pone_descripcio = "Ventas"
       Case "G"
            pone_descripcio = "OtraVenta"
       Case "E"
            pone_descripcio = "NotaCreV"
       Case "F"
            pone_descripcio = "NotaDebV"
            
       Case "J", "K", "L", "M"
            pone_descripcio = "Compras"
       Case "P"
            pone_descripcio = "OtraCompra"
       Case "N"
            pone_descripcio = "NotaCreC"
       Case "O"
            pone_descripcio = "NotaDebC"
       Case "V"
            pone_descripcio = "R.Egreso"
       Case "W"
            pone_descripcio = "R.Ingreso"
End Select
End Function
Function pone_fpago(buf As String) As String
Select Case buf
       Case "A"
            pone_fpago = "Soles"
       Case "B"
            pone_fpago = "Dolares"
       Case "C"
            pone_fpago = "Credito"
       Case "D"
            pone_fpago = "T.Credito"
       Case "F"
            pone_fpago = "T.Debito"
       Case "G"
            pone_fpago = "Letra"
       Case "H"
            pone_fpago = "Vales"
       Case "K"
            pone_fpago = "AntBco"
       Case "I"
            pone_fpago = "Otros"
End Select

End Function

