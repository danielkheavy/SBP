VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form expmovca 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimiento de Caja"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   14865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   14865
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
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
      Height          =   7335
      Left            =   120
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   14655
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
         TabIndex        =   48
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
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
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
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "expmovca.frx":0000
         Height          =   6375
         Left            =   120
         OleObjectBlob   =   "expmovca.frx":0014
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   840
         Width           =   14415
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
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8280
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Mensajes"
      Height          =   5415
      Left            =   2640
      TabIndex        =   41
      Top             =   1440
      Visible         =   0   'False
      Width           =   9015
      Begin VB.Label Label4 
         Caption         =   "ESPERE UN MOMENTO..PROCESANDO..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   480
         TabIndex        =   42
         Top             =   720
         Width           =   7935
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "expmovca.frx":09DF
      Height          =   6735
      Left            =   120
      OleObjectBlob   =   "expmovca.frx":09F3
      TabIndex        =   30
      Top             =   1320
      Width           =   14535
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   14160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFF80&
      Height          =   1275
      Left            =   120
      ScaleHeight     =   1215
      ScaleWidth      =   14700
      TabIndex        =   0
      Top             =   0
      Width           =   14760
      Begin VB.TextBox nombre 
         Height          =   375
         Left            =   2520
         MaxLength       =   11
         TabIndex        =   43
         Text            =   "%"
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox local1 
         Height          =   375
         Left            =   5280
         MaxLength       =   3
         TabIndex        =   28
         Text            =   "%"
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox codigo 
         Height          =   375
         Left            =   2520
         MaxLength       =   11
         TabIndex        =   26
         Text            =   "%"
         Top             =   120
         Width           =   1815
      End
      Begin VB.ComboBox ordenado 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10920
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox fpago 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8040
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
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
         Left            =   13200
         MaskColor       =   &H00FFFFFF&
         Picture         =   "expmovca.frx":2C6A
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.ComboBox tipo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10920
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox fechai 
         Height          =   375
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   12
         Top             =   120
         Width           =   1815
      End
      Begin VB.TextBox fechaf 
         Height          =   375
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   11
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox cajero 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8040
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox caja 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8040
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   120
         Width           =   1815
      End
      Begin VB.ComboBox turno 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10920
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   120
         Width           =   1815
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
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "expmovca.frx":3418
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
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
         Left            =   0
         MaskColor       =   &H00E0E0E0&
         Picture         =   "expmovca.frx":462A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprimir"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         Height          =   375
         Left            =   1560
         TabIndex        =   44
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
         Height          =   375
         Left            =   4320
         TabIndex        =   29
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         Height          =   375
         Left            =   1560
         TabIndex        =   27
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ordenado Por"
         Height          =   375
         Left            =   9840
         TabIndex        =   21
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FormaPago"
         Height          =   375
         Left            =   7080
         TabIndex        =   19
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoDocto"
         Height          =   375
         Left            =   9840
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   375
         Left            =   4320
         TabIndex        =   14
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   375
         Left            =   4320
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Turno"
         Height          =   375
         Left            =   9840
         TabIndex        =   10
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja"
         Height          =   375
         Left            =   7080
         TabIndex        =   9
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cajero"
         Height          =   375
         Left            =   7080
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "US$"
      Height          =   375
      Left            =   4680
      TabIndex        =   40
      Top             =   8520
      Width           =   975
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S/."
      Height          =   375
      Left            =   4680
      TabIndex        =   39
      Top             =   8160
      Width           =   975
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cargo"
      Height          =   375
      Left            =   5640
      TabIndex        =   38
      Top             =   8520
      Width           =   975
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo"
      Height          =   375
      Left            =   10560
      TabIndex        =   37
      Top             =   8520
      Width           =   975
   End
   Begin VB.Label cargod 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6600
      TabIndex        =   36
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Abono"
      Height          =   375
      Left            =   7920
      TabIndex        =   35
      Top             =   8520
      Width           =   975
   End
   Begin VB.Label abonod 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8880
      TabIndex        =   34
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Label saldod 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11520
      TabIndex        =   33
      Top             =   8520
      Width           =   1455
   End
   Begin VB.Label saldos 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11520
      TabIndex        =   32
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Label abonos 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8880
      TabIndex        =   31
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Abono"
      Height          =   375
      Left            =   7920
      TabIndex        =   25
      Top             =   8160
      Width           =   975
   End
   Begin VB.Label cargos 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6600
      TabIndex        =   24
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo"
      Height          =   375
      Left            =   10560
      TabIndex        =   23
      Top             =   8160
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cargo"
      Height          =   375
      Left            =   5640
      TabIndex        =   22
      Top             =   8160
      Width           =   975
   End
   Begin VB.Label afecta 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   13920
      TabIndex        =   4
      Top             =   7440
      Width           =   735
   End
   Begin VB.Label acu 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   13680
      TabIndex        =   3
      Top             =   7440
      Width           =   255
   End
   Begin VB.Menu dki222 
      Caption         =   "&Ejecuta"
   End
   Begin VB.Menu dki232312 
      Caption         =   "&Imprime"
   End
   Begin VB.Menu lfo3434 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "expmovca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancelar_Click()

End Sub

Private Sub cmdDelete_Click()
dbo912_Click
End Sub

Private Sub cmdGrabar_Click()
End Sub

Private Sub cmdSort_Click()

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   lfo3434_Click
   Exit Sub
End If
Command2_Click

End Sub

Private Sub cmdExit_Click()
lfo3434_Click
End Sub

Private Sub cmdPrint_Click()
Dim found As Integer
If Frame1.Visible = True Then Exit Sub

Filename = globaldir & "\temporal\" & gusuario & ".txt"
    borrar_archivo Filename
    cerrar_archivo
    found = borra_nombre("" & Filename)
    Open Filename For Append As #1
    '------------------------------------
    cabecera_documento1
    cuerpo_programa_documento1
    '------------------------------------
    Close #1
    cerrar_archivo
    genver.File = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Sub consulta_codigo()
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.ListIndex = 0
Frame2.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "1"
Command2_Click

End Sub

Private Sub codigo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
consulta_codigo
End If

End Sub

Private Sub Command1_Click()
If Frame1.Visible = True Then Exit Sub
Frame1.Visible = True
xborrar
sql_recibos
Frame1.Visible = False
End Sub




Private Sub dbo912_Click()

End Sub

Private Sub dki9923_Click()
End Sub

Private Sub dnu823_Click()
End Sub

Private Sub Command2_Click()
ejecuta 1
End Sub

Private Sub DBGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
If opcion1 = "1" Then
   codigo = dbGrid2.Columns(1)
   Frame2.Visible = False
   codigo.SetFocus
   'codigo_KeyPress 13
End If
End If

End Sub

Private Sub DBGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
Dim buf As String
Dim buf2 As String
Dim sw As Integer
If KeyCode <> 13 And KeyCode <> 27 Then
          'MsgBox KeyCode
          If KeyCode >= 48 And KeyCode <= 57 Then
             GoTo sigue9
          End If
          If KeyCode >= 65 And KeyCode <= 90 Then
             GoTo sigue9
          End If
          If KeyCode >= 97 And KeyCode <= 122 Then
             GoTo sigue9
          End If
          If KeyCode = 8 And Chr(KeyCode) = "*" Then
             GoTo sigue9
          End If
          Exit Sub
sigue9:
          If KeyCode = 8 Then
            If Len(buffer) > 0 Then
               buf = Mid$(buffer, 1, Len(buffer) - 1)
               buffer = buf
               KeyCode = 0
               Else
               KeyCode = 0
               Exit Sub
            End If
         End If
         buf = Chr(KeyCode)
         If Chr(KeyCode) = "*" Then
            buf = ""
            buffer = buf
         End If
         If KeyCode <> 13 Then
            buffer = buffer + buf
         End If
         buf = buffer
         ejecuta 0
End If

End Sub

Private Sub dki222_Click()
If Frame2.Visible = True Then Exit Sub

If Frame1.Visible = True Then Exit Sub
Command1_Click
End Sub

Private Sub dki232312_Click()
If Frame2.Visible = True Then Exit Sub

If Frame1.Visible = True Then Exit Sub
cmdPrint_Click
End Sub

Private Sub Form_Activate()
fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
fechaf = Format(Now, "dd/mm/yyyy")
carga_inicial

End Sub
Sub carga_inicial()
Dim mytablex As Table
'rubro.Clear
'rubro.AddItem "ReciboCaja"
'rubro.AddItem "*"
'rubro.ListIndex = 0

cajero.Clear
cajero.AddItem "%"
Set mytablex = mydbxglo.OpenTable("vendedor")
Do
If mytablex.EOF Then Exit Do
cajero.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
mytablex.MoveNext
Loop
mytablex.Close
cajero.ListIndex = 0

caja.Clear
caja.AddItem "%"
Set mytablex = mydbxglo.OpenTable("parameca")
Do
If mytablex.EOF Then Exit Do
If "" & mytablex.Fields("terminal") = "C" Then
caja.AddItem "" & mytablex.Fields("caja") & "|" & mytablex.Fields("DESCRIPCIO")
End If
mytablex.MoveNext
Loop
mytablex.Close
caja.ListIndex = 0

turno.Clear
turno.AddItem "%"
Set mytablex = mydbxglo.OpenTable("turno")
Do
If mytablex.EOF Then Exit Do
turno.AddItem "" & mytablex.Fields("turno") & "|" & mytablex.Fields("DESCRIPCIO")
mytablex.MoveNext
Loop
mytablex.Close
turno.ListIndex = 0


tipo.Clear
tipo.AddItem "%"
Set mytablex = mydbxglo.OpenTable("tipo")
Do
If mytablex.EOF Then Exit Do
If "" & mytablex.Fields("tipodoc") = "W" Or "" & mytablex.Fields("tipodoc") = "V" Then
   tipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("DESCRIPCIO")
End If
mytablex.MoveNext
Loop
mytablex.Close
tipo.ListIndex = 0

fpago.Clear
fpago.AddItem "%"
Set mytablex = mydbxglo.OpenTable("fpago")
Do
If mytablex.EOF Then Exit Do
fpago.AddItem "" & mytablex.Fields("fpago") & "|" & mytablex.Fields("DESCRIPCIO")
mytablex.MoveNext
Loop
mytablex.Close
fpago.ListIndex = 0


End Sub

Private Sub Form_Load()
ordenado.Clear
ordenado.AddItem "fecha"
ordenado.AddItem "tipo"
ordenado.AddItem "val(numero)"
ordenado.AddItem "Codigo"
ordenado.AddItem "Usuario"
ordenado.AddItem "caja"
ordenado.AddItem "turno"
ordenado.AddItem "fpago"
ordenado.AddItem "orden"
ordenado.AddItem "observa"
ordenado.AddItem "descripcio"
ordenado.AddItem "nombre"
ordenado.ListIndex = 0

Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.ListIndex = 0
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)
If Frame1.Visible = True Then
   Frame1.Visible = False
End If
End Sub

Private Sub lfo3434_Click()
If Frame2.Visible = True Then
   Frame2.Visible = False
   Exit Sub
End If

If Frame1.Visible = True Then
   Frame1.Visible = False
   Exit Sub
End If
expmovca.Hide
Unload expmovca
End Sub
Sub sql_recibos()
On Error GoTo cmd37_err
Dim vr
Dim found As Integer
Dim buf As String
Dim mytableY As Table
Dim mytablex As Snapshot
If Not IsDate(fechai) Then Exit Sub
If Not IsDate(fechaf) Then Exit Sub
Set mytableY = mydbxglo.OpenTable("_b" + gusuario)
buf = "select * from fpagov where "
buf = buf & "  fecha>=" & "DateValue('" & fechai & "'" & ")"
buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
If local1 <> "%" Then
   buf = buf & " and local='" & local1 & "'"
End If
If cajero <> "%" Then
   buf = buf & " and usuario='" & extra_loquesea(cajero) & "'"
End If
If tipo <> "%" Then
   buf = buf & " and tipo='" & extra_loquesea(tipo) & "'"
End If
If caja <> "%" Then
   buf = buf & " and caja='" & extra_loquesea(caja) & "'"
End If
If turno <> "%" Then
   buf = buf & " and turno='" & extra_loquesea(turno) & "'"
End If
If fpago <> "%" Then
   buf = buf & " and fpago='" & extra_loquesea(fpago) & "'"
End If
If codigo <> "%" Then
buf = buf & " and codigo like '" & codigo & "'"
End If
If Nombre <> "%" Then
buf = buf & " and nombre like '" & Nombre & "'"
End If
'buf = buf & " and tipoclie='C' "
buf = buf & " and estado='2' "
buf = buf & " and (acu='W' or acu='V')"
'If rubro = "ReciboCaja" Then
'   buf = buf & " and (acu='W' or acu='V') "
'End If
'If rubro <> "ReciboCaja" Then
   'buf = buf & " and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='V' or acu='W' or acu='J' or acu='K' or  acu='L' or acu='M' or acu='P' ) "
'End If
buf = buf & " order by " & ordenado & ",str(numero)"
Set mytablex = mydbxglo.CreateSnapshot(buf)
Do
If mytablex.EOF Then Exit Do
vr = DoEvents()
If Frame1.Visible = False Then
   Exit Do
End If
mytableY.AddNew
mytableY.Fields("local") = "" & mytablex.Fields("local")
mytableY.Fields("observa") = busca_observa(mytablex)
mytableY.Fields("tipo") = "" & mytablex.Fields("tipo")
mytableY.Fields("tipoclie") = "" & mytablex.Fields("tipoclie")
mytableY.Fields("nombret") = busca_tipo("" & mytablex.Fields("tipo"))
mytableY.Fields("serie") = "" & mytablex.Fields("serie")
mytableY.Fields("numero") = "" & mytablex.Fields("numero")
mytableY.Fields("codigo") = "" & mytablex.Fields("codigo")
mytableY.Fields("nombre") = "" & mytablex.Fields("nombre")
mytableY.Fields("fecha") = "" & mytablex.Fields("fecha")
mytableY.Fields("moneda") = "" & mytablex.Fields("moneda")
mytableY.Fields("usuario") = "" & mytablex.Fields("usuario")
mytableY.Fields("caja") = "" & mytablex.Fields("caja")
mytableY.Fields("turno") = "" & mytablex.Fields("turno")
mytableY.Fields("fpago") = "" & mytablex.Fields("descripcio")
If "" & mytablex.Fields("acu") = "A" Or "" & mytablex.Fields("acu") = "B" Or "" & mytablex.Fields("acu") = "C" Or "" & mytablex.Fields("acu") = "D" Or "" & mytablex.Fields("acu") = "G" Then
   If "" & mytablex.Fields("acufp") = "C" Then
   mytableY.Fields("acu") = "C"
   mytableY.Fields("cargo") = Val("" & mytablex.Fields("total"))
   End If
   If "" & mytablex.Fields("acufp") <> "C" Then
   mytableY.Fields("acu") = "A"
   mytableY.Fields("abono") = Val("" & mytablex.Fields("total"))
   End If
End If
If "" & mytablex.Fields("acu") = "J" Or "" & mytablex.Fields("acu") = "K" Or "" & mytablex.Fields("acu") = "L" Or "" & mytablex.Fields("acu") = "M" Or "" & mytablex.Fields("acu") = "P" Then
   If "" & mytablex.Fields("acufp") = "C" Then
   mytableY.Fields("acu") = "A"
   mytableY.Fields("abono") = Val("" & mytablex.Fields("total"))
   End If
   If "" & mytablex.Fields("acufp") <> "C" Then
   mytableY.Fields("acu") = "C"
   mytableY.Fields("cargo") = Val("" & mytablex.Fields("total"))
   End If
End If

If "" & mytablex.Fields("acu") = "W" Then 'ingreso
   mytableY.Fields("acu") = "A"
   mytableY.Fields("abono") = Val("" & mytablex.Fields("total"))
End If
If "" & mytablex.Fields("acu") = "V" Then 'ingreso
   mytableY.Fields("acu") = "C"
   mytableY.Fields("cargo") = Val("" & mytablex.Fields("total"))
End If

mytableY.Update
sigamos:
mytablex.MoveNext
Loop
mytablex.Close
mytableY.Close
'xborrar

               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldir
               Data2.RecordSource = "select * from _b" & gusuario
               Data2.Refresh
               
               sumar_recibos
               'DBGrid2.SetFocus
               'MsgBox "xx"
               
Frame1.Visible = False
Exit Sub
cmd37_err:
Frame1.Visible = False
MsgBox "Error en select " & error$, 48, "Aviso"
Exit Sub

End Sub
Sub xborrar()
On Error GoTo cmd112_err
Data2.Database.Execute "DELETE FROM _b" & gusuario
Exit Sub
cmd112_err:
Exit Sub

End Sub
Sub sumar_recibos()
Dim xcargos As Double
Dim xabonos As Double
Dim xcargod As Double
Dim xabonod As Double

xcargos = 0
xabonos = 0
xcargod = 0
xabonod = 0

Data2.Refresh
Do
If Data2.Recordset.EOF Then Exit Do
If "" & Data2.Recordset.Fields("moneda") = "S" Then
xcargos = xcargos + Val("" & Data2.Recordset.Fields("cargo"))
xabonos = xabonos + Val("" & Data2.Recordset.Fields("abono"))
End If
If "" & Data2.Recordset.Fields("moneda") = "D" Then
xcargod = xcargod + Val("" & Data2.Recordset.Fields("cargo"))
xabonod = xabonod + Val("" & Data2.Recordset.Fields("abono"))
End If

Data2.Recordset.MoveNext
Loop
cargos = Format(xcargos, "0.00")
abonos = Format(xabonos, "0.00")
saldos = Format(xcargos - xabonos, "0.00")
cargod = Format(xcargod, "0.00")
abonod = Format(xabonod, "0.00")
saldod = Format(xcargod - xabonod, "0.00")

End Sub
Function busca_tipo(buf As String)
Dim mytablex As Table
   Set mytablex = mydbxglo.OpenTable("tipo")
   mytablex.Index = "tipo"
   mytablex.Seek "=", buf
   If Not mytablex.NoMatch Then
      busca_tipo = "" & mytablex.Fields("descripcio")
   End If
   mytablex.Close

End Function
Sub cabecera_documento1()
Dim buf As String
Dim i As Integer
Dim found As Integer
    If contlin > 0 Then
       buf = Chr$(12)
       found = formateaa(buf, Len(buf), 0, 0)
    End If
    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = "Movimiento de Caja  "
    found = formateaa(buf, 90, 2, 0)
    
    found = formateaa("Lo", 3, 0, 0)
    found = formateaa("Tp", 3, 0, 0)
    found = formateaa("Srie", 5, 0, 0)
    found = formateaa("Numero", 12, 0, 0)
    found = formateaa("Fecha", 11, 0, 0)
    found = formateaa("Nombre", 31, 0, 0)
    
    found = formateaa("Cargo ", 11, 0, 1)
    found = formateaa("Abono ", 11, 0, 1)
    found = formateaa("Saldo", 11, 2, 1)
        
    buf = String(150, "-")
    found = formateaa(buf, 150, 2, 0)
    

End Sub
Sub cuerpo_programa_documento1()
Dim buf As String
Dim found As Integer
Dim xcargo As Double
Dim xabono As Double
On Error GoTo cmd78812_err
      xcargo = 0
      xabono = 0
Data2.Refresh
Do
If Data2.Recordset.EOF Then Exit Do
      buf = "" & Data2.Recordset.Fields("LOCAL")
      found = formateaa(buf, 2, 0, 0)
      found = formateaa("", 1, 0, 0)
      buf = "" & Data2.Recordset.Fields("tipo")
      found = formateaa(buf, 2, 0, 0)
      found = formateaa("", 1, 0, 0)
      buf = "" & Data2.Recordset.Fields("serie")
      found = formateaa(buf, 4, 0, 0)
      found = formateaa("", 1, 0, 0)
      buf = "" & Data2.Recordset.Fields("numero")
      found = formateaa(buf, 11, 0, 0)
      found = formateaa("", 1, 0, 0)
      buf = "" & Data2.Recordset.Fields("fecha")
      found = formateaa(buf, 10, 0, 0)
      found = formateaa("", 1, 0, 0)
      buf = "" & Data2.Recordset.Fields("nombre")
      found = formateaa(buf, 30, 0, 0)
      found = formateaa("", 1, 0, 0)
      buf = Format(Val("" & Data2.Recordset.Fields("cargo")), "0.00")
      found = formateaa(buf, 10, 0, 1)
      found = formateaa("", 1, 0, 0)
      buf = Format(Val("" & Data2.Recordset.Fields("abono")), "0.00")
      found = formateaa(buf, 10, 0, 1)
      found = formateaa("", 1, 2, 0)
      
      
      nlineas
      
      
      xcargo = xcargo + Val("" & Data2.Recordset.Fields("cargo"))
      xabono = xabono + Val("" & Data2.Recordset.Fields("abono"))
      
      Data2.Recordset.MoveNext
Loop

      found = formateaa("", 65, 0, 0)
      buf = Format(xcargo, "0.00")
      found = formateaa(buf, 10, 0, 1)
      found = formateaa("", 1, 0, 0)
      buf = Format(xabono, "0.00")
      found = formateaa(buf, 10, 0, 1)
      found = formateaa("", 1, 0, 0)
      buf = Format(xcargo - xabono, "0.00")
      found = formateaa(buf, 10, 0, 1)
      found = formateaa("", 1, 2, 0)
      
Exit Sub
cmd78812_err:
Exit Sub
End Sub
Sub nlineas()
    contlin = contlin + 1
    If contlin > 45 Then
       cabecera_documento1
    End If
End Sub
Function escajachica(buf As String)
Dim mytablex As Table
   Set mytablex = mydbxglo.OpenTable("tipo")
   mytablex.Index = "tipo"
   mytablex.Seek "=", buf
   If Not mytablex.NoMatch Then
      If "" & mytablex.Fields("cajachica") = "C" Then
         escajachica = 1
      End If
   End If
   mytablex.Close

End Function
Function busca_observa(mytableY As Table) As String
Dim mytablex As Table
   Set mytablex = mydbxglo.OpenTable("recibo")
   mytablex.Index = "recibo"
   mytablex.Seek "=", "" & mytableY.Fields("local"), "" & mytableY.Fields("tipo"), "" & mytableY.Fields("serie"), "" & mytableY.Fields("numero")
   If Not mytablex.NoMatch Then
   busca_observa = "" & mytablex.Fields("observa")
   End If
   mytablex.Close
End Function
Sub ejecuta(sw As Integer)
Dim buf As String
   If opcion1 = "1" Then
      If Len(buffer) = 0 Then
      buf = "select Nombre,Codigo from clientes"
      Else
      buf = "select Nombre,Codigo from clientes where  " & Combo1 & " like '" & buffer & "%'"
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
               If opcion1 = "20" Then
                  dbGrid2.Columns(0).Width = 2000
                  dbGrid2.Columns(1).Width = 1300
               End If
               If sw = 1 Then
                  dbGrid2.SetFocus
               End If
End Sub




